using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

namespace BusinessCats
{
    public class SeriousBusinessCat
    {
        protected MainWindow _main;

        public WhisperCat whisperCat;

        Microsoft.Lync.Model.LyncClient client = null;
        Microsoft.Lync.Model.Extensibility.Automation automation = null;

        string _selfUserName = "";
        string _selfUserNameLast = "";

        private List<LyncConversation> _conversations = new List<LyncConversation>();


        public SeriousBusinessCat(MainWindow main)
        {
            this._main = main;

            this.whisperCat = new WhisperCat(this);

            try
            {
#if DEBUG
                var _testWhisper = new WhisperWindow(null, null, null, "Your key: 12345\r\nTheir key: 6789A");
     
                _testWhisper.AddWhisper(new Whisper(true, "testing a really really long message, at least it seems pretty long, but i guess it is really not that long to begin with", "cipher1"));
                _testWhisper.AddWhisper(new Whisper(false, "okay, looks good", "cipher2"));
                _testWhisper.AddWhisper(new Whisper(true, "glad you think so, you jerk", "cipher3"));
                _testWhisper.Show();
#endif


                //Start the conversation
                automation = LyncClient.GetAutomation();
                client = LyncClient.GetClient();

                ConversationManager conversationManager = client.ConversationManager;
                conversationManager.ConversationAdded += ConversationManager_ConversationAdded;
                conversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;
                //conversationManager.ConversationAdded += new EventHandler<ConversationManagerEventArgs>(conversationManager_ConversationAdded);

                _main.lbConversations.ItemsSource = _conversations;

                RefreshConversations(true);

                try
                {
                    _selfUserName = client.Self.Contact.GetContactInformation(ContactInformationType.PrimaryEmailAddress).ToString();
                    if (_selfUserName.Contains("@"))
                    {
                        _selfUserName = _selfUserName.Split('@')[0];
                    }
                    _selfUserNameLast = _selfUserName;
                    if (_selfUserNameLast.Contains('.'))
                    {
                        _selfUserNameLast = _selfUserNameLast.Split('.')[1];
                    }
                }
                catch (Exception)
                {
                    // ignore
                }
            }
            catch (LyncClientException lyncClientException)
            {
                MessageBox.Show("Failed to connect to Lync: " + lyncClientException.ToString());
                Console.Out.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (Util.IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    MessageBox.Show("Failed to connect to Lync with system error: " + systemException.ToString());
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }


        private void ConversationManager_ConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            RefreshConversations();
        }

        private void ConversationManager_ConversationAdded(object sender, ConversationManagerEventArgs e)
        {
            RefreshConversations();
        }

        public LyncConversation FindConversationFromTitle(string title)
        {
            try {
                var matches = new List<LyncConversation>();

                if (title.StartsWith("Conversation ("))
                {
                    string titleCountString = Regex.Match(title, @"\d+").Value;
                    int count = 0;
                    if (!int.TryParse(titleCountString, out count))
                    {
                        return null;
                    }
                    // try to find a conversation with correct number of participants
                    foreach (var c in _conversations)
                    {
                        if (c.conversation.Participants.Count == count)
                        {
                            matches.Add(c);
                        }
                    }
                }
                else
                {
                    // will be the name of the contact
                    foreach (var c in _conversations)
                    {
                        if (c.conversation.Participants.Count > 2)
                        {
                            continue;
                        }

                        foreach (var p in c.conversation.Participants)
                        {
                            if (p.IsSelf)
                            {
                                continue;
                            }

                            object val;
                            if (p.Properties.TryGetValue(ParticipantProperty.Name, out val) && val is string)
                            {
                                string name = (string)val;
                                if (name == title)
                                {
                                    matches.Add(c);
                                }
                            }
                        }
                    }
                }

                if (matches.Count == 0)
                {
                    return null;
                }

                if (matches.Count == 1)
                {
                    return matches[0];
                }

                matches = matches.OrderByDescending(c => {
                    object val;
                    if (c.conversation.Properties.TryGetValue(ConversationProperty.LastActivityTimeStamp, out val) && val is DateTime)
                    {
                        var lastActivity = (DateTime)val;
                        return lastActivity;
                    }
                    else {
                        return DateTime.MinValue;
                    }
                }).ToList();

                return matches[0];
            } catch (Exception)
            {
                // ignore it
                return null;
            }
        }

        public LyncConversation FindConversation(Conversation conv)
        {
            foreach (var c in _conversations)
            {
                if (c.conversation == conv)
                {
                    return c;
                }
            }
            return null;
        }

        public LyncConversation TrackConversation(Conversation conv)
        {
            var participants = new Dictionary<Participant, LyncParticipant>();
            foreach (var participant in conv.Participants)
            {
                if (participant.IsSelf)
                {
                    continue;
                }

                InstantMessageModality modality = (InstantMessageModality)participant.Modalities[ModalityTypes.InstantMessage];
                modality.InstantMessageReceived += imModality_InstantMessageReceived;

                participants.Add(participant, new LyncParticipant { modality = modality, participant = participant });
            }

            conv.ParticipantAdded += Conv_ParticipantAdded;
            conv.ParticipantRemoved += Conv_ParticipantRemoved;

            var lyncConv = new LyncConversation() { conversation = conv, participants = participants };

            lyncConv.UpdateDesc();

            return lyncConv;
        }

        public void UpdateConversation(LyncConversation conv)
        {
            var activeParticipants = new List<Participant>();
            foreach (var participant in conv.conversation.Participants)
            {
                if (participant.IsSelf)
                {
                    continue;
                }

                activeParticipants.Add(participant);

                if (conv.participants.ContainsKey(participant))
                {
                    continue;
                }

                InstantMessageModality modality = (InstantMessageModality)participant.Modalities[ModalityTypes.InstantMessage];
                modality.InstantMessageReceived += imModality_InstantMessageReceived;

                conv.participants.Add(participant, new LyncParticipant { modality = modality, participant = participant });
            }

            var deadParticipants = new List<Participant>();
            foreach (var participant in conv.participants.Keys)
            {
                if (!activeParticipants.Contains(participant))
                {
                    deadParticipants.Add(participant);
                }
            }

            foreach (var participant in deadParticipants)
            {
                conv.participants.Remove(participant);
            }

            conv.UpdateDesc();
        }

        private void Conv_ParticipantRemoved(object sender, ParticipantCollectionChangedEventArgs e)
        {
            RefreshConversations();
        }

        private void Conv_ParticipantAdded(object sender, ParticipantCollectionChangedEventArgs e)
        {
            RefreshConversations();
        }

        protected void DoRefreshConversations(bool selectActive = false)
        {
            _conversations.RemoveAll((c) => {
                if (c.conversation.State == ConversationState.Terminated)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            });

            foreach (var lyncConv in client.ConversationManager.Conversations)
            {
                var conv = FindConversation(lyncConv);
                if (null == conv)
                {
                    _conversations.Add(TrackConversation(lyncConv));
                    continue;
                }
                else
                {
                    UpdateConversation(conv);
                }
            }

            _conversations.Sort(Comparer<LyncConversation>.Create((l, r) => {
                int lGroup = l.participants.Count > 2 ? 1 : 0;
                int rGroup = r.participants.Count > 2 ? 1 : 0;

                if (lGroup > rGroup)
                {
                    return -1;
                }
                else if (lGroup < rGroup)
                {
                    return 1;
                }
                else {
                    return string.Compare(l.ToString(), r.ToString());
                }
            }));

            _main.lbConversations.Items.Refresh();

            if (selectActive)
            {
                string activeConversationTitle = Util.GetActiveLyncConversationTitle();
                var activeConversation = FindConversationFromTitle(activeConversationTitle);
                if (activeConversation != null)
                {
                    _main.lbConversations.SelectedItem = activeConversation;
                }
            }

            //var selfContact = client.Self.Contact;

            //foreach (var infoType in Enum.GetValues(typeof(ContactInformationType)).Cast<ContactInformationType>())
            //{
            //    try {
            //        var infoTypeName = Enum.GetName(typeof(ContactInformationType), infoType);

            //        var infoValue = selfContact.GetContactInformation(infoType).ToString();

            //        if (!string.IsNullOrEmpty(infoValue))
            //        {
            //            System.Diagnostics.Debug.Print($"{infoTypeName}: {infoValue}");
            //        }
            //    } catch (Exception)
            //    {
            //        // ignore
            //    }
            //}
        }

        public void RefreshConversations(bool selectActive = false)
        {
            _main.lbConversations.Dispatcher.Invoke(() => { DoRefreshConversations(selectActive); });
        }

        void imModality_InstantMessageReceived(object sender, MessageSentEventArgs e)
        {
            try
            {
                var modality = (InstantMessageModality)sender;
                IDictionary<InstantMessageContentType, string> messageFormatProperty = e.Contents;

                string source = (string)modality.Participant.Contact.GetContactInformation(ContactInformationType.DisplayName);

                string text = "";
                string raw = "";
                if (messageFormatProperty.ContainsKey(InstantMessageContentType.PlainText))
                {
                    if (messageFormatProperty.TryGetValue(InstantMessageContentType.PlainText, out raw))
                    {
                        text = raw;
                    }
                }
                else if (messageFormatProperty.ContainsKey(InstantMessageContentType.RichText))
                {
                    if (messageFormatProperty.TryGetValue(InstantMessageContentType.RichText, out raw))
                    {
                        text = Util.TextFromRtf(raw);
                    }
                }
                else if (messageFormatProperty.ContainsKey(InstantMessageContentType.Html))
                {
                    if (messageFormatProperty.TryGetValue(InstantMessageContentType.Html, out raw))
                    {

                        var reg = new System.Text.RegularExpressions.Regex("<[^>]+>");
                        var stripped = reg.Replace(raw, "");
                        text = System.Web.HttpUtility.HtmlDecode(stripped);
                    }
                }
                //else if (messageFormatProperty.ContainsKey(InstantMessageContentType.Gif))
                //{
                //    if (messageFormatProperty.TryGetValue(InstantMessageContentType.Gif, out raw))
                //    {
                //        // gif, "base64:{data}"
                //    }
                //}
                //else if (messageFormatProperty.ContainsKey(InstantMessageContentType.Ink))
                //{
                //    // InkSerializedFormat (application/x-ms-ink)
                //    if (messageFormatProperty.TryGetValue(InstantMessageContentType.Ink, out raw))
                //    {
                //        // isf, just use Microsoft.Ink to Load/Save
                //    }
                //}

                if (!string.IsNullOrEmpty(text))
                {
                    if (whisperCat.HandleWhisper(modality.Conversation, modality.Participant, text))
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void StartWhispering()
        {
            var conv = GetActiveConversation();
            if (conv == null)
            {
                return;
            }

            if (conv.participants.Count > 1)
            {
                MessageBox.Show("It is hard to make friends in a crowded hallway.");
                return;
            }

            var participant = conv.participants.Values.First();

            whisperCat.InitiateWhisperHandshake(conv.conversation, participant.participant);
        }


        private LyncConversation GetActiveConversation()
        {
            if (_main.lbConversations.SelectedItem == null)
            {
                MessageBox.Show("Please select a business conversation with your business associates to whom you intend to send your business cat ascii art", "whoa there partner");
                return null;
            }
            var conv = (LyncConversation)_main.lbConversations.SelectedItem;
            return conv;
        }

        public void DoSendImage(string data)
        {
            if (string.IsNullOrEmpty(data))
            {
                return;
            }

            var conv = GetActiveConversation();
            if (conv == null)
            {
                return;
            }

            DoSendMessage(conv.conversation, data, InstantMessageContentType.Gif);
        }

        void DoSendText(string data, InstantMessageContentType type)
        {
            if (string.IsNullOrEmpty(data))
            {
                return;
            }

            var conv = GetActiveConversation();
            if (conv == null)
            {
                return;
            }

            DoSendMessage(conv.conversation, data, type);
        }

        protected async Task DoSendMessageAsync(Conversation conversation, string data, InstantMessageContentType type)
        {
            try
            {
                InstantMessageModality imModality = (InstantMessageModality)conversation.Modalities[ModalityTypes.InstantMessage];
                var formattedMessage = new Dictionary<InstantMessageContentType, string>();
                formattedMessage.Add(type, data);

                if (imModality.CanInvoke(ModalityAction.SendInstantMessage))
                {
                    await Task.Factory.FromAsync(imModality.BeginSendMessage, imModality.EndSendMessage, formattedMessage, imModality);
                }
            }
            catch (LyncClientException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void DoSendMessage(Conversation conversation, string data, InstantMessageContentType type)
        {
            var task = DoSendMessageAsync(conversation, data, type);
            //task.Wait(100);
        }

        public void Blp()
        {
            string data = @"{\rtf\ansi lol}";
            DoSendText(data, InstantMessageContentType.RichText);
        }

        protected string ImageToData(System.Drawing.Image image)
        {
            using (var memStream = new System.IO.MemoryStream())
            {
                var quantizer = new ImageQuantization.OctreeQuantizer(255, 8);
                using (var quantized = quantizer.Quantize(image))
                {
                    quantized.Save(memStream, System.Drawing.Imaging.ImageFormat.Gif);
                }

                //image.Save(memStream, System.Drawing.Imaging.ImageFormat.Gif);

                return "base64:" + System.Convert.ToBase64String(memStream.ToArray());
            }
        }
        protected string ScaleToWidth(System.Drawing.Image image, int width)
        {
            var newWidth = width;
            var newHeight = (int)((image.Height * (width / (double)image.Width)) + 0.5);

            string data = "";

            if (newWidth == image.Width && newHeight == image.Height)
            {
                data = ImageToData(image);
            }
            else {
                using (var newImage = new System.Drawing.Bitmap(image, newWidth, newHeight))
                {
                    data = ImageToData(newImage);
                }
            }

            System.Diagnostics.Debug.Print("Resize ({0},{1}) -> ({2},{3}) -- len={4}", image.Width, image.Height, newWidth, newHeight, data.Length);

            return data;
        }

        protected const int threshold = 65000;

        public string DataFromImage(System.Drawing.Image image, bool useMaxWidth = true)
        {

            int maxWidth = useMaxWidth ? _main.CalcMaxWidth() : int.MaxValue;
            if (maxWidth <= 0)
            {
                maxWidth = 1;
            }

            if (maxWidth > image.Width)
            {
                maxWidth = image.Width;
            }
            
            // The quantizer runs super slow with big images so start at a small reasonable size
            int newWidth = 640; // starting point

            if (newWidth > maxWidth)
            {
                newWidth = maxWidth;
            }

            string data = "";

            // keep scaling down until we go below the threshold
            data = ScaleToWidth(image, newWidth);

            if (data.Length < threshold && newWidth < maxWidth)
            {
                // already below threshold at our starting point? try starting at 1600 px then

                newWidth = 1600;

                if (newWidth > maxWidth)
                {
                    newWidth = maxWidth;
                }

                data = ScaleToWidth(image, newWidth);
            }

            if (data.Length < threshold && newWidth < maxWidth)
            {
                // still below threshold at our starting point? start with the real size of the image i suppose, up to the hardlimit
                const int hardLimit = 3840;

                newWidth = image.Width;
                if (newWidth > hardLimit)
                {
                    newWidth = hardLimit;
                }

                if (newWidth > maxWidth)
                {
                    newWidth = maxWidth;
                }

                data = ScaleToWidth(image, newWidth);
            }

            while (data.Length >= threshold)
            {
                int nextWidth = newWidth;

                // scale down
                double diffRatio = Math.Sqrt(threshold / (double)data.Length);
                nextWidth = (int)(nextWidth * diffRatio);

                // align to 8 px
                nextWidth = (nextWidth / 8) * 8;

                if (nextWidth == newWidth)
                {
                    nextWidth -= 8;
                }

                if (nextWidth <= 0)
                {
                    throw new Exception("Something very troubling is going on resizing this image you gave me.");
                }

                newWidth = nextWidth;
                data = ScaleToWidth(image, newWidth);
            }

            // now we are below the threshold! this time let's just move up until we go over.
            while (newWidth < maxWidth)
            {
                newWidth += 16;
                string newData = ScaleToWidth(image, newWidth);
                if (newData.Length < threshold)
                {
                    data = newData;
                } else
                {
                    break;
                }
            }

            return data;
        }

        public string DataFromFile(string filename)
        {
            using (var image = System.Drawing.Image.FromFile(filename))
            {
                if (System.IO.Path.GetExtension(filename).ToLower() == "gif" && image.Width < _main.CalcMaxWidth())
                {
                    string data = "base64:" + System.Convert.ToBase64String(System.IO.File.ReadAllBytes(filename));
                    if (data.Length < threshold)
                    {
                        return data;
                    }
                }

                return DataFromImage(image);
            }
        }

        public string DataFromClipboard()
        {
            using (var image = System.Windows.Forms.Clipboard.GetImage())
            {
                if (image == null)
                {
                    return "";
                }
                return DataFromImage(image);
            }
        }

        public string DataFromString(string data)
        {
            using (var stream = new System.IO.MemoryStream(System.Convert.FromBase64String(data)))
            {
                using (var image = System.Drawing.Image.FromStream(stream))
                {
                    return DataFromImage(image);
                }
            }
        }

        private void SaveSnip(System.Drawing.Image bmp)
        {
            var snipPath = Util.GetSnipPath();
            var fileName = Util.GetUtcTimeStamp() + ".png";

            var filePath = System.IO.Path.Combine(snipPath, fileName);

            bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
        }

        public void Snip()
        {
            _main.Visibility = Visibility.Hidden;
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
            {
                try
                {
                    using (var bmp = SnippingTool.Snip())
                    {
                        _main.Visibility = Visibility.Visible;
                        if (bmp != null)
                        {
                            SaveSnip(bmp);
                            System.Windows.Forms.Clipboard.SetImage(bmp);
                        }
                    }
                }
                finally
                {
                    _main.Visibility = Visibility.Visible;
                }
            }));
        }

        public void SnipAndSend()
        {
            _main.Visibility = Visibility.Hidden;
            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
            {
                try
                {
                    using (var bmp = SnippingTool.Snip())
                    {
                        _main.Visibility = Visibility.Visible;
                        if (bmp != null)
                        {
                            SaveSnip(bmp);
                            DoSendImage(DataFromImage(bmp, useMaxWidth: false));

                            try
                            {
                                System.Windows.Forms.Clipboard.SetImage(bmp);
                            }
                            catch (Exception)
                            {
                                // ignore
                            }
                        }
                    }
                }
                finally
                {
                    _main.Visibility = Visibility.Visible;
                }
            }));
        }

        public void Paste()
        {
            if (System.Windows.Forms.Clipboard.ContainsImage())
            {
                DoSendImage(DataFromClipboard());
            }
            else if (System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Rtf))
            {
                string rtf = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Rtf);
                string data = rtf.Replace(@"{\rtf\", @"{\rtf1\");
                DoSendText(data, InstantMessageContentType.RichText);
            }
            //else if (System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Html))
            //{
            //    DoSendText(System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Html), InstantMessageContentType.Html);
            //}
            else if (System.Windows.Forms.Clipboard.ContainsText())
            {
                DoSendText(System.Windows.Forms.Clipboard.GetText(), InstantMessageContentType.PlainText);
            }
        }
    }
}
