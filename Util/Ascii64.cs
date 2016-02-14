using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessCats
{
    class Ascii64
    {
        /// <summary>
        /// simple Decode with default values
        /// </summary>
        public static byte[] Decode(string data)
        {
            return Convert.FromBase64String(data);
        }

        /// <summary>
        /// simple Encode with default values
        /// </summary>
        public static string Encode(byte[] data)
        {
            return Convert.ToBase64String(data);
        }
    }
}
