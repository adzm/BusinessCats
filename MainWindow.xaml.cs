using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Documents;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using MessageBox = System.Windows.MessageBox;
using System.Configuration;
using System.Security.Cryptography;

using MahApps.Metro;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Windows.Threading;
using System.IO;
using System.Diagnostics;

namespace BusinessCats
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public SeriousBusinessCat _seriousBusiness;

        private string _hotkeys = "";

        public MainWindow()
        {
            InitializeComponent();

            this.MetroDialogOptions.ColorScheme = MetroDialogColorScheme.Accented;

            this.Closing += (s, e) =>
            {
                try
                {
                    if (_seriousBusiness?.whisperCat != null)
                    {
                        _seriousBusiness.whisperCat.CloseAll();
                    }
                }catch(Exception) {
                    // ignore
                }
            };

            this.PreviewKeyDown += (s, e) =>
            {
                const string hotkeyCheck = "X\x58\x5AZW\x59\x57Ym\x6C\x46";

                char keyCode = (char)(((int)e.Key) + 64);

                _hotkeys += keyCode;

                if (_hotkeys.EndsWith(hotkeyCheck, StringComparison.Ordinal))
                {
                    _hotkeys = "";
                    e.Handled = true;

                    BusinessCats.Properties.Settings.Default.KrahsWobniar = !BusinessCats.Properties.Settings.Default.KrahsWobniar;                    
                    krahsWobniar.Visibility = CanKrahsWobniarAllNight;

                    BusinessCats.Properties.Settings.Default.Save();

                    return;
                }
                else
                {
                    while (_hotkeys.Length > 0 && !hotkeyCheck.StartsWith(_hotkeys))
                    {
                        _hotkeys = _hotkeys.Substring(1);
                    }

                    if (_hotkeys.Length >= hotkeyCheck.Length - 3)
                    {
                        if (keyCode == 'm' || keyCode == 'l')
                        {
                            e.Handled = true;
                        }
                    }
                }
            };

            _seriousBusiness = new SeriousBusinessCat(this);
        }

        private void btnKrahsWobniar_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.DoSendImage(_seriousBusiness.DataFromString(WellKnownCats.GetKrahsWobniarData()));
        }

        private void btnPaste_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.Paste();
        }

        private void btnFlash_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.Blp();
        }

        public int CalcMaxWidth()
        {
            int value = (int)sliderBig.Value;
            int max = (int)sliderBig.Maximum;
            int width = (int)sliderBig.ActualWidth;

            if (value == max)
            {
                return int.MaxValue;
            }

            int calcWidth = (width * value) / max;

            if (calcWidth < 16)
            {
                calcWidth = 16;
            }

            return calcWidth;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                if (files.Length == 0) {
                    return;
                }

                string data = _seriousBusiness.DataFromFile(files[0]);
                _seriousBusiness.DoSendImage(data);
            }
        }


        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.RefreshConversations(true);
        }

        private void btnSnip_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.Snip();
        }
        private void btnSnipAndSend_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.SnipAndSend();
        }

        public bool EatenByLions
        {
            get
            {
                return BusinessCats.Properties.Settings.Default.EatenByLions;
            }
            set
            {
                BusinessCats.Properties.Settings.Default.EatenByLions = value;

                BusinessCats.Properties.Settings.Default.Save();
            }
        }

        public Visibility CanKrahsWobniarAllNight
        {
            get
            {
                return BusinessCats.Properties.Settings.Default.KrahsWobniar ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        public bool IsDarkTheme
        {
            get
            {
                return BusinessCats.Properties.Settings.Default.ThemeName == "BaseDark";
            }
            set
            {
                string themeName = "";
                if (value)
                {
                    themeName = "BaseDark";
                }
                else
                {
                    themeName = "BaseLight";
                }

                if (themeName != BusinessCats.Properties.Settings.Default.ThemeName) {
                    BusinessCats.Properties.Settings.Default.ThemeName = themeName;

                    BusinessCats.Properties.Settings.Default.Save();

                    ApplyTheme();
                }
            }
        }

        private void ApplyTheme()
        {
            var themeName = BusinessCats.Properties.Settings.Default.ThemeName;

            var appTheme = ThemeManager.GetAppTheme(themeName);
            if (appTheme != null)
            {
                ThemeManager.ChangeAppStyle(Application.Current,
                                            ThemeManager.GetAccent("Blue"),
                                            appTheme);
            }
        }

        private void btnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(Util.GetStableUserPath());
        }        

        private void btnSecret_Click(object sender, RoutedEventArgs e)
        {
            _seriousBusiness.StartWhispering();
        }
    }
}
