using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using MahApps.Metro.Controls;
using MahApps.Metro;

namespace BusinessCats
{
    /// <summary>
    /// Interaction logic for WhisperWindow.xaml
    /// </summary>
    public partial class WhisperWindow : MetroWindow
    {
        public class Whisper
        {
            public string Text { get; set; }

            public bool IsSelf { get; set; }

            public Whisper()
            { }

            public Whisper(bool isSelf, string text)
            {
                IsSelf = isSelf;
                Text = text;
            }
        }

        public ICommand SendMessageCmd { get; set; }

        protected void InitCommands()
        {
            SendMessageCmd = new CommandCat(() => {
                string text = tbMessage.Text;

                if (!seriousBusiness.whisperCat.SendWhisper(conversation, participant, text))
                {
                    return;
                }

                _whispers.Add(new Whisper(true, text));

                TextBoxHelper.SetWatermark(tbMessage, "...");
                tbMessage.Text = "";

                var border = (Border)VisualTreeHelper.GetChild(lbWhispers, 0);
                var scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                scrollViewer.ScrollToBottom();
            });
        }

        public WhisperWindow(Conversation conversation, Participant participant, SeriousBusinessCat seriousBusiness)
        {
            this.conversation = conversation;
            this.participant = participant;
            this.seriousBusiness = seriousBusiness;

            InitializeComponent();

            this.DataContext = this;

            InitCommands();

            this.Activated += (s,e) => { FlashWindow.Stop(this); };

//#if DEBUG
//            _whispers.Add(new Whisper(true, "testing a really really long message, at least it seems pretty long, but i guess it is really not that long to begin with"));
//            _whispers.Add(new Whisper(false, "Okay, looks good"));
//            _whispers.Add(new Whisper(true, "Glad you think so, you jerk"));
//#endif

            lbWhispers.ItemsSource = _whispers;

            string bob = "neighbor cat";

            if (participant != null)
            {
                bob = participant.Contact.GetContactInformation(ContactInformationType.DisplayName).ToString();
            }

            textBlock.Text = bob;

            this.Title = $"Whispering with {bob}";

            FocusManager.SetFocusedElement(this, tbMessage);
        }        

        public Conversation conversation;
        public Participant participant;
        public WhisperCat whisperCat;
        public SeriousBusinessCat seriousBusiness;

        private ObservableCollection<Whisper> _whispers = new ObservableCollection<Whisper>();

        public void AddWhisper(string text)
        {
            this.Dispatcher.Invoke(() => {
                _whispers.Add(new Whisper(false, text));
                var border = (Border)VisualTreeHelper.GetChild(lbWhispers, 0);
                var scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                scrollViewer.ScrollToBottom();

                if (!this.IsActive)
                {
                    FlashWindow.Flash(this);
                }
            });
        }
    }
}
