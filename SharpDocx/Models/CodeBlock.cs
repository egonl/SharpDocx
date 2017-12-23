using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.Models
{
    public class CodeBlock
    {
        public string Code { get; set; }

        public Text Placeholder { get; set; }

        public int CurlyBracketLevelIncrement => GetCurlyBracketLevelIncrement(Code);

        internal Text StartText, EndText;

        public CodeBlock(string code)
        {
            Code = code;
        }

        internal void RemoveEmptyParagraphs()
        {
            var startParagraph = StartText.GetParent<Paragraph>();
            if (startParagraph?.Parent != null && !startParagraph.HasText())
            {
                startParagraph.Remove();
            }

            var endParagraph = EndText.GetParent<Paragraph>();
            if (endParagraph?.Parent != null && !endParagraph.HasText())
            {
                endParagraph.Remove();
            }
        }

        public static int GetCurlyBracketLevelIncrement(string code)
        {
            // Note: since this isn't a proper C# parser, this code won't work properly when there are curly brackets in comments/strings/etc.
            var increment = 0;

            foreach (var c in code)
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

        public static string GetExpression(string code, char startExpression, char endExpression, int startIndex = 0)
        {
            // Note: same issue as above.
            startIndex = code.IndexOf(startExpression, startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var expression = new StringBuilder();
            var increment = 0;
            for (var i = startIndex; i < code.Length; ++i)
            {
                expression.Append(code[i]);

                if (code[i] == startExpression)
                {
                    ++increment;
                }
                else if (code[i] == endExpression)
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

        public static string GetExpressionInBrackets(string code, int startIndex = 0)
        {
            return GetExpression(code, '(', ')', startIndex);
        }

        public static string GetExpression(string code, string startOrEndTag, int startIndex = 0, bool removeStartAndEndTag = true)
        {
            // The start/end tag can be escaped with '\'.
            code = code.Replace("\\" + startOrEndTag, "b260231509a148aa9d751a5a9d79abd7");

            startIndex = code.IndexOf(startOrEndTag, startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var endIndex = code.IndexOf(startOrEndTag, startIndex + startOrEndTag.Length);
            if (endIndex == -1)
            {
                return null;
            }

            code = removeStartAndEndTag 
                ? code.Substring(startIndex + 1, endIndex - startIndex - 1) 
                : code.Substring(startIndex, endIndex - startIndex + 1);

            return code.Replace("b260231509a148aa9d751a5a9d79abd7", "\\" + startOrEndTag);
        }

        public static string GetExpressionInApostrophes(string code, int startIndex = 0, bool removeApostrophes = true)
        {
            return GetExpression(code, "\"", startIndex, removeApostrophes);
        }
    }
}