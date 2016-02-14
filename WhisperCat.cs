using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace BusinessCats
{
    public class WhisperCat
    {
        public SeriousBusinessCat seriousBusiness;

        public WhisperCat(SeriousBusinessCat seriousBusiness)
        {
            this.seriousBusiness = seriousBusiness;
        }

        public static string ShhWrap(string data)
        {
            string rtf = Util.RtfFromText(data);

            if (rtf.StartsWith(@"{\rtf1"))
            {
                rtf = rtf.Remove(5, 1);
            }

            {
                string unrtfData = Util.TextFromRtf(rtf);

                Debug.Assert(unrtfData == data);
            }

            return rtf;
        }

        readonly static string whisperHandshakePrefix = "psst....";
        readonly static string whisperPrefix = "shhh....";

        public bool HandleWhisper(Conversation conversation, Participant participant, string text)
        {
            if (text.StartsWith(whisperHandshakePrefix) && text.StartsWith(whisperHandshakePrefix + "<~"))
            {
                return CompleteWhisperHandshake(conversation, participant, Ascii85.Decode(text.Substring(whisperHandshakePrefix.Length)));
            }
            else if (text.StartsWith(whisperPrefix) && text.StartsWith(whisperPrefix + "<~"))
            {
                return ExtractWhisper(conversation, participant, Ascii85.Decode(text.Substring(whisperHandshakePrefix.Length)));
            }

            return false;
        }

        private void SendWhisperData(Conversation conversation, byte[] data)
        {
            var whisperText = whisperPrefix + Ascii85.Encode(data);

            bool sendAsRtf = true;

            if (sendAsRtf)
            {
                seriousBusiness.DoSendMessage(conversation, ShhWrap(whisperText), InstantMessageContentType.RichText);
            }
            else {
                seriousBusiness.DoSendMessage(conversation, whisperText, InstantMessageContentType.PlainText);
            }
        }

        private void SendWhisperHandshake(Conversation conversation, byte[] data)
        {
            var handshakeText = whisperHandshakePrefix + Ascii85.Encode(data);

            bool sendAsRtf = false;

            if (sendAsRtf)
            {
                seriousBusiness.DoSendMessage(conversation, ShhWrap(handshakeText), InstantMessageContentType.RichText);
            }
            else {
                seriousBusiness.DoSendMessage(conversation, handshakeText, InstantMessageContentType.PlainText);
            }
        }

        public bool InitiateWhisperHandshake(Conversation conversation, Participant participant)
        {
            var conv = seriousBusiness.FindConversation(conversation);
            if (conv == null)
            {
                return false;
            }

            var p = conv.FindParticipant(participant);
            if (p == null)
            {
                return false;
            }

            var pubKey = p.GetPublicKey();

            SendWhisperHandshake(conv.conversation, pubKey);

            return true;
        }

        private bool CompleteWhisperHandshake(Conversation conversation, Participant participant, byte[] keyBlob)
        {
            var conv = seriousBusiness.FindConversation(conversation);
            if (conv == null)
            {
                return false;
            }

            var p = conv.FindParticipant(participant);
            if (p == null)
            {
                return false;
            }

            if (p.dh == null)
            {
                if (!InitiateWhisperHandshake(conversation, participant))
                {
                    return false;
                }
            }

            p.DeriveKey(keyBlob);

            Application.Current.Dispatcher.Invoke(() => { FindOrCreateWhisperWindow(conversation, participant).Show(); });

            return true;
        }

        private bool ExtractWhisper(Conversation conversation, Participant participant, byte[] textBlob)
        {
            var conv = seriousBusiness.FindConversation(conversation);
            if (conv == null)
            {
                return false;
            }

            var p = conv.FindParticipant(participant);
            if (p == null)
            {
                return false;
            }

            if (p.dh == null)
            {
                //InitiateWhisperHandshake(conversation, participant);
                return false;
            }

            if (p.derivedKey == null)
            {
                return false;
            }

            byte[] iv = new byte[16];
            for (int i = 0; i < 16; ++i)
            {
                iv[i] = textBlob[i];
            }

            byte[] plainText;
            using (Aes aes = new AesCryptoServiceProvider())
            {
                aes.Key = p.derivedKey;
                aes.IV = iv;

                using (var plainTextStream = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(plainTextStream, aes.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(textBlob, 16, textBlob.Length - 16);
                        cs.Close();
                        plainText = plainTextStream.ToArray();
                    }
                }
            }

            string plainTextString = Encoding.UTF8.GetString(plainText);

            FindWhisperWindow(conversation, participant).AddWhisper(plainTextString);

            return true;
        }

        public bool SendWhisper(Conversation conversation, Participant participant, string text)
        {
            var conv = seriousBusiness.FindConversation(conversation);
            if (conv == null)
            {
                return false;
            }

            var p = conv.FindParticipant(participant);
            if (p == null)
            {
                return false;
            }

            if (p.dh == null)
            {
                return false;
            }

            if (p.derivedKey == null)
            {
                return false;
            }

            var bytes = Encoding.UTF8.GetBytes(text);

            byte[] cipherText;
            using (Aes aes = new AesCryptoServiceProvider())
            {
                aes.Key = p.derivedKey;
                byte[] iv = aes.IV;

                if (iv.Length != 16)
                {
                    throw new Exception("Expecting 256-bit iv");
                }

                using (var cipherTextStream = new MemoryStream())
                {
                    cipherTextStream.Write(iv, 0, iv.Length);

                    using (CryptoStream cs = new CryptoStream(cipherTextStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytes, 0, bytes.Length);
                        cs.Close();
                        cipherText = cipherTextStream.ToArray();
                    }
                }
            }

            SendWhisperData(conv.conversation, cipherText);

            return true;
        }


        private List<WhisperWindow> _whisperWindows = new List<WhisperWindow>();

        public WhisperWindow FindWhisperWindow(Conversation conversation, Participant participant)
        {
            foreach (var window in _whisperWindows)
            {
                if (window.conversation == conversation && window.participant == participant)
                {
                    return window;
                }
            }

            return null;
        }

        public WhisperWindow FindOrCreateWhisperWindow(Conversation conversation, Participant participant)
        {
            WhisperWindow window = FindWhisperWindow(conversation, participant);

            if (window == null)
            {
                window = new WhisperWindow(conversation, participant, seriousBusiness);
                _whisperWindows.Add(window);

                window.Closed += (s, e) => { _whisperWindows.Remove(window); };
                window.Show();
            }

            return window;
        }

        public void CloseAll()
        {
            // copy, since the Closing event mutates _whisperWindows
            var whisperWindows = new List<WhisperWindow>(_whisperWindows);
            foreach(var window in whisperWindows)
            {
                window.Close();
            }
        }
    }
}
