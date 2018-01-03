using System.Text;

namespace SharpDocx.Extensions
{
    internal static class StringExtensions
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

        public static int GetCurlyBracketLevelIncrement(this string s)
        {
            // Note: since this isn't a proper C# parser, this code won't work properly when there are curly brackets in comments/strings/etc.
            var increment = 0;

            foreach (var c in s)
            {
                if (c == '{')
                {
                    ++increment;
                }
                else if (c == '}')
                {
                    --increment;
                }
            }

            return increment;
        }

        public static string GetExpression(this string s, char startExpression, char endExpression, int startIndex = 0)
        {
            // Note: same issue as above.
            startIndex = s.IndexOf(startExpression, startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var expression = new StringBuilder();
            var increment = 0;
            for (var i = startIndex; i < s.Length; ++i)
            {
                expression.Append(s[i]);

                if (s[i] == startExpression)
                {
                    ++increment;
                }
                else if (s[i] == endExpression)
                {
                    --increment;
                }

                if (increment == 0)
                {
                    return expression.ToString();
                }
            }

            return null;
        }

        public static string GetExpressionInBrackets(this string s, int startIndex = 0)
        {
            return s.GetExpression('(', ')', startIndex);
        }

        public static string GetExpression(this string s, string startOrEndTag, int startIndex = 0,
            bool removeStartAndEndTag = true)
        {
            // The start/end tag can be escaped with '\'.
            s = s.Replace("\\" + startOrEndTag, "b260231509a148aa9d751a5a9d79abd7");

            startIndex = s.IndexOf(startOrEndTag, startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var endIndex = s.IndexOf(startOrEndTag, startIndex + startOrEndTag.Length);
            if (endIndex == -1)
            {
                return null;
            }

            s = removeStartAndEndTag
                ? s.Substring(startIndex + 1, endIndex - startIndex - 1)
                : s.Substring(startIndex, endIndex - startIndex + 1);

            return s.Replace("b260231509a148aa9d751a5a9d79abd7", "\\" + startOrEndTag);
        }

        public static string GetExpressionInApostrophes(this string s, int startIndex = 0,
            bool removeApostrophes = true)
        {
            return s.GetExpression("\"", startIndex, removeApostrophes);
        }
    }
}