using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDocx.Models
{
    public class CodeBlock : TextPart
    {
        public string Code { get; set; }

        public Text Placeholder { get; set; }

        public bool Conditional { get; set; }

        public string Condition { get; set; }

        public Text EndConditionalPart { get; set; }

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
    }
}