using System.Text;

namespace SharpDocx.Extensions
{
    public static class StringExtensions
    {
        public static string RemoveRedundantWhitespace(this string s)
        {
            var sb = new StringBuilder();
            var previousCharIsWhitespace = false;

            foreach (var c in s)
            {
                if (char.IsWhiteSpace(c) && previousCharIsWhitespace)
                {
                    continue;
                }

                previousCharIsWhitespace = char.IsWhiteSpace(c);
                sb.Append(c);
            }

            return sb.ToString();
        }
    }
}