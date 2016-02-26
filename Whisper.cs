using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessCats
{
    public class Whisper
    {
        public string Text { get; set; }

        public string CipherText { get; set; }

        public bool IsSelf { get; set; }

        public Whisper()
        { }

        public Whisper(bool isSelf, string text, string cipherText)
        {
            IsSelf = isSelf;
            Text = text;
            CipherText = cipherText;
        }
    }
}
