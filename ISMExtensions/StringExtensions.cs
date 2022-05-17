using Newtonsoft.Json;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ISMExtensions.Extensions
{
    public static class StringExtensions
    {
        public static string JoinString(this string[] arr, string joinCharacter = ",")
        {
            string r = "";
            if (arr == null)
                return r;
            foreach (var item in arr)
            {
                r += item;
                if (item != arr.Last())
                    r += joinCharacter;
            }
            return r;
        }

        public static IEnumerable<string> Splitx(this string str, int chunkSize)
        {
            return Enumerable.Range(0, str.Length / chunkSize)
                .Select(i => str.Substring(i * chunkSize, chunkSize));
        }

        public static string Replace(this string str, string[] oldValueList, string newvalue)
        {
            foreach (var v in oldValueList)
            {
                str = str.Replace(v, newvalue);
            }
            return str;
        }

        public static bool IsDigit(this string str)
        {
            foreach (var c in str)
            {
                if (!char.IsDigit(c))
                    return false;
            }
            return true;
        }

        public static string ToLowerCulture(this string str)
        {
            if (string.IsNullOrEmpty(str))
                return string.Empty;
            return str.ToLower(new CultureInfo("tr-TR", false));
        }

        public static string ToUpperFirstCharacter(this string str)
        {
            if (!string.IsNullOrEmpty(str))
                return str.Length > 1 ? str[0].ToString().ToUpperInvariant() + str.Substring(1, str.Length - 1) : str;
            return string.Empty;
        }

        public static bool IsEmail(this string str)
        {
            const string MatchEmailPattern =
                  @"^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"
           + @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?
				[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
           + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?
				[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"
           + @"([a-zA-Z0-9]+[\w-]+\.)+[a-zA-Z]{1}[a-zA-Z0-9-]{1,23})$";

            if (str != null) return Regex.IsMatch(str, MatchEmailPattern);
            else return false;
        }

        public static string Join(this List<string> lst, string seperator)
        {
            if (lst == null || lst.Count == 0)
                return string.Empty;

            string t = "";

            for (int i = 0; i < lst.Count; i++)
            {
                t += lst[i];
                if (i != lst.Count - 1)
                    t += ", ";
            }

            return t;
        }

        public static T ToObject<T>(this string value) where T : class
        {
            return string.IsNullOrEmpty(value) ? null : JsonConvert.DeserializeObject<T>(value);
        }

        public static string GetPlainTextFromHtml(this string value)
        {
            var htmlString = value;

            if (string.IsNullOrEmpty(htmlString))
                return htmlString;

            string htmlTagPattern = "<.*?>";
            htmlString = Regex.Replace(htmlString, htmlTagPattern, string.Empty);
            htmlString = Regex.Replace(htmlString, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&ccedil;", "ç", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&ouml;", "ö", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&uuml;", "ü", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&Ccedil;", "Ç", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&Ouml;", "Ö", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&Uuml;", "Ü", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&rdquo;", "”", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&ldquo;", "“", RegexOptions.Multiline);
            htmlString = Regex.Replace(htmlString, @"&rsquo;", "’", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&nbsp;", " ", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&#39;", "'", RegexOptions.Multiline);

            htmlString = Regex.Replace(htmlString, @"&quot;", "\"", RegexOptions.Multiline);

            return htmlString.Trim();
        }
    }
}
