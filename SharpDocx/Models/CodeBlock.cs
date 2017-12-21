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

        public Text StartText, EndText;

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

        public static string GetExpressionInBrackets(string code, int startIndex = 0)
        {
            // Note: same issue as above.
            startIndex = code.IndexOf("(", startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var expression = new StringBuilder();
            var increment = 0;
            for (var i = startIndex; i < code.Length; ++i)
            {
                expression.Append(code[i]);

                if (code[i] == '(')
                {
                    ++increment;
                }
                else if (code[i] == ')')
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
    }
}