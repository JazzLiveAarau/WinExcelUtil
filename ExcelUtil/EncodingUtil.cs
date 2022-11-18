using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtil
{
    /// <summary>Utility functions for encoding</summary>
    public class EncodingUtil
    {

        /// <summary>Convert from Unicode to ASCII</summary>
        public static void UnicodeToAscii(string i_unicode_string, out string o_ascii_string)
        {
            // http://msdn.microsoft.com/en-us/library/system.text.encoding(v=vs.71).aspx

            // http://msdn.microsoft.com/en-us/library/system.text.encoding.aspx
            // Encoding is the process of transforming a set of Unicode characters into a sequence of bytes
            // Decoding is the process of transforming a sequence of encoded bytes into a set of Unicode characters

            //QQstring i_unicode_string = "This string contains the unicode character Pi(\u03a0)";

            // Create two different encodings.
            Encoding ascii = Encoding.ASCII;
            Encoding unicode = Encoding.Unicode;

            // Convert the string into a byte[].
            byte[] unicode_bytes = unicode.GetBytes(i_unicode_string);

            // Perform the conversion from one encoding to the other.
            byte[] ascii_bytes = Encoding.Convert(unicode, ascii, unicode_bytes);

            // Convert the new byte[] into a char[] and then into a string.
            // This is a slightly different approach to converting to illustrate
            // the use of GetCharCount/GetChars.
            char[] ascii_chars = new char[ascii.GetCharCount(ascii_bytes, 0, ascii_bytes.Length)];
            ascii.GetChars(ascii_bytes, 0, ascii_bytes.Length, ascii_chars, 0);
            o_ascii_string = new string(ascii_chars);
        }
    }
}
