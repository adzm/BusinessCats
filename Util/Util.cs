using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BusinessCats
{
    public static class Util
    {
        /// <summary>
        /// Identify if a particular SystemException is one of the exceptions which may be thrown
        /// by the Lync Model API.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        public static bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }

        public static string GetTimeStamp(DateTime dt)
        {
            return dt.ToString("s").Replace("-", "").Replace(":", "") + "." + dt.Millisecond.ToString("D3");
        }
        public static string GetUtcTimeStamp()
        {
            return GetTimeStamp(DateTime.UtcNow);
        }
        public static string GetLocalTimeStamp()
        {
            return GetTimeStamp(DateTime.Now);
        }

        public static string GetConfigPath()
        {
            string path = System.IO.Path.GetDirectoryName(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath);
            System.IO.Directory.CreateDirectory(path);
            return path;
        }

        public static string GetStableUserPath(params string[] subdirectories)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData, Environment.SpecialFolderOption.Create);
            path = System.IO.Path.Combine(path, "BusinessCats");

            foreach(var subdirectory in subdirectories)
            {
                path = System.IO.Path.Combine(path, subdirectory);
            }

            System.IO.Directory.CreateDirectory(path);

            return path;
        }

        public static string GetSnipPath()
        {
            return GetStableUserPath("Snips");
        }

        public static string GetArchivePath(string subdirectory = "")
        {
            return GetStableUserPath("Archive", subdirectory);
        }

        public static string TextFromRtf(string rtf)
        {
            using (var rtb = new System.Windows.Forms.RichTextBox())
            {
                rtb.Rtf = rtf;
                return rtb.Text;
            }
        }

        public static string RtfFromText(string text)
        {
            using (var rtb = new System.Windows.Forms.RichTextBox())
            {
                rtb.Text = text;
                return rtb.Rtf;
            }
        }

        [DllImport("user32.dll")]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, [Out] StringBuilder lParam);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        const int WM_GETTEXT = 0x000D;
        const int WM_GETTEXTLENGTH = 0x000E;

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        public static string GetWindowTextRaw(IntPtr hwnd)
        {
            StringBuilder sb = new StringBuilder(256);
            SendMessage(hwnd, WM_GETTEXT, (IntPtr)sb.Capacity, sb);
            return sb.ToString();
        }

        public static IntPtr GetActiveLyncConversationWnd()
        {
            IntPtr activeConversationHandle = IntPtr.Zero;
            EnumWindows((hWnd, lParam) => {
                var className = new StringBuilder(256);

                if (0 != GetClassName(hWnd, className, className.Capacity))
                {
                    if (0 == string.Compare("LyncConversationWindowClass", className.ToString(), false, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        activeConversationHandle = hWnd;
                        return false;
                    }
                }

                return true;
            }, IntPtr.Zero);

            return activeConversationHandle;
        }

        public static string GetActiveLyncConversationTitle()
        {
            IntPtr activeWnd = GetActiveLyncConversationWnd();

            if (activeWnd != IntPtr.Zero)
            {
                return GetWindowTextRaw(activeWnd);
            } else
            {
                return "";
            }
        }
    }
}
