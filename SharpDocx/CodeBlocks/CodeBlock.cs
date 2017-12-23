using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.CodeBlocks
{
    public class CodeBlock
    {
        public string Code { get; }

        public Text Placeholder { get; internal set; }

        internal Text StartText { get; set; }

        internal Text EndText { get; set; }

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
    }
}