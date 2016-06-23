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
        public ICommand SendMessageCmd { get; set; } 

        protected void InitCommands()
        {
            SendMessageCmd = new CommandCat(() => {
                string text = tbMessage.Text;

                string cipherText = seriousBusiness.whisperCat.SendWhisper(conversation, participant, text);

                if (string.IsNullOrEmpty(cipherText))
                {
                    return;
                }

                TextBoxHelper.SetWatermark(tbMessage, "...");
                tbMessage.Text = "";

                AddWhisper(new Whisper(true, text, cipherText));
            });
        }

        public WhisperWindow(Conversation conversation, Participant participant, SeriousBusinessCat seriousBusiness, string info)
        {
            this.conversation = conversation;
            this.participant = participant;
            this.seriousBusiness = seriousBusiness;

            InitializeComponent();

            this.DataContext = this;

            InitCommands();

            this.Activated += (s,e) => { FlashWindow.Stop(this); };
            
            lbWhispers.ItemsSource = _whispers;

            string bob = "neighbor cat";

            if (participant != null)
            {
                bob = participant.Contact.GetContactInformation(ContactInformationType.DisplayName).ToString();
            }

            textBlock.Text = bob;

            if (!string.IsNullOrEmpty(info))
            {
                textBlock.ToolTip = info;
            }

            this.Title = $"Whispering with {bob}";

            FocusManager.SetFocusedElement(this, tbMessage);
        }        

        public Conversation conversation;
        public Participant participant;
        public WhisperCat whisperCat;
        public SeriousBusinessCat seriousBusiness;

        private ObservableCollection<Whisper> _whispers = new ObservableCollection<Whisper>();

        public void AddWhisper(Whisper whisper)
        {
            this.Dispatcher.Invoke(() => {
                _whispers.Add(whisper);

                if (IsLoaded)
                {
                    var border = (Border)VisualTreeHelper.GetChild(lbWhispers, 0);
                    var scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                    scrollViewer.ScrollToBottom();

                    if (!whisper.IsSelf && !this.IsActive)
                    {
                        FlashWindow.Flash(this);
                    }
                }
            });
        }
    }
}
