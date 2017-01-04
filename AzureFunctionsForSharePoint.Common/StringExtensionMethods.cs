using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Handy string extensions for parsing text
    /// </summary>
    public static class StringExtensionMethods
    {
        /// <summary>
        /// Returns the text between the first case sensitive instances of a given preamble and postscript or an empty string.
        /// </summary>
        /// <param name="input">The string to parse</param>
        /// <param name="preamble">The beginning string to match</param>
        /// <param name="postscript">The ending string to match</param>
        /// <returns>The string between the preamble and postscript or nothing</returns>
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

        /// <summary>
        /// Returns the text between the first case insensitive instances of a given preamble and postscript or an empty string.
        /// </summary>
        /// <param name="input">The string to parse</param>
        /// <param name="preamble">The beginning string to match</param>
        /// <param name="postscript">The ending string to match</param>
        /// <param name="ignoreCase">Ignore the case of the preamble and postscript if true</param>
        /// <returns>The string between the preamble and postscript or nothing</returns>
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

        /// <summary>
        /// Returns a list of strings from a string between the instances of given preamble and postscript pairs or an empty list.
        /// </summary>
        /// <param name="input">The string to parse</param>
        /// <param name="preamble">The beginning string to match</param>
        /// <param name="postscript">The ending string to match</param>
        /// <param name="ignoreCase">Ignore the case of the preamble and postscript if true</param>
        /// <returns>A list of strings</returns>
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
