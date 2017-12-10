using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.Models
{
    public class CodeBlock : MapPart
    {
        public string Code { get; set; }

        public Text Placeholder { get; set; }

        public bool Conditional { get; set; }

        public string Condition { get; set; }

        public Text EndConditionalPart { get; set; }

        public Text StartText, EndText;

        public int CurlyBracketLevelIncrement
        {
            get
            {
                // Note: since this isn't a proper C# parser, this code won't work properly when there are curly brackets in comments/strings/etc.
                var increment = 0;

                foreach (var c in Code)
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
        }

        public string GetExpressionInBrackets(int startIndex = 0)
        {
            // Note: same issue as above.
            startIndex = Code.IndexOf("(", startIndex);
            if (startIndex == -1)
            {
                return null;
            }

            var expression = new StringBuilder();
            var increment = 0;
            for (var i = startIndex; i < Code.Length; ++i)
            {
                expression.Append(Code[i]);

                if (Code[i] == '(')
                {
                    ++increment;
                }
                else if (Code[i] == ')')
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

        public void RemoveEmptyParagraphs()
        {
            var startParagraph = this.StartText.GetParent<Paragraph>();
            if (startParagraph != null && startParagraph.Parent != null && !startParagraph.HasText())
            {
                startParagraph.Remove();
            }

            var endParagraph = this.EndText.GetParent<Paragraph>();
            if (endParagraph != null && endParagraph.Parent != null && !endParagraph.HasText())
            {
                endParagraph.Remove();
            }
        }
    }
}