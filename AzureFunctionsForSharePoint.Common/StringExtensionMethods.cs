using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctionsForSharePoint.Common
{
    public static class StringExtensionMethods
    {
        public static string GetInnerText(this string input, string preamble, string postscript)
        {
            var retVal = string.Empty;

            if (input.Contains(preamble))
            {
                var start = input.IndexOf(preamble, StringComparison.Ordinal) + preamble.Length;
                var end = input.IndexOf(postscript, start, StringComparison.Ordinal);
                if (end == -1) return string.Empty;
                retVal = input.Substring(start, end - start);
            }
            return retVal;
        }

        public static string GetInnerText(this string input, string preamble, string postscript, bool ignoreCase)
        {
            if (!ignoreCase) return input.GetInnerText(preamble, postscript);

            var retVal = string.Empty;

            if (input.Contains(preamble))
            {
                var start = input.IndexOf(preamble, StringComparison.InvariantCultureIgnoreCase) + preamble.Length;
                var end = input.IndexOf(postscript, start, StringComparison.InvariantCultureIgnoreCase);
                if (end == -1) return string.Empty;
                retVal = input.Substring(start, end - start);
            }
            return retVal;
        }

        public static byte[] ToByteArrayUtf8(this string input)
        {
            return Encoding.UTF8.GetBytes(input);
        }

        public static List<string> GetInnerTextList(this string input, string preamble, string postscript)
        {
            return input.GetInnerTextList(preamble, postscript, false);
        }

        public static List<string> GetInnerTextList(this string input, string preamble, string postscript,
            bool ignoreCase)
        {
            var retVal = new List<string>();

            string lastVal;
            var current = input;
            do
            {
                lastVal = current.GetInnerText(preamble, postscript, ignoreCase);
                if (lastVal != string.Empty)
                {
                    if (!retVal.Contains(lastVal)) retVal.Add(lastVal);
                    current = current.Substring(current.IndexOf(lastVal, StringComparison.Ordinal) + lastVal.Length - 1);
                }
            } while (lastVal != "");

            return retVal;
        }
    }
}
