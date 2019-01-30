using System.Linq;
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

        internal InsertionPoint CurrentInsertionPoint { get; set; }

        private CodeBlock()
        {
            CurrentInsertionPoint = null;
        }

        internal CodeBlock(string code)
        {
            Code = code;
        }

        internal void RemoveEmptyParagraphs()
        {
            var startParagraph = StartText.GetParent<Paragraph>();
            if (startParagraph?.Parent != null && 
                !startParagraph.HasText() && 
                CanDeleteParagraph(startParagraph))
            {
                startParagraph.Remove();
            }

            var endParagraph = EndText.GetParent<Paragraph>();
            if (endParagraph?.Parent != null && 
                !endParagraph.HasText() &&
                CanDeleteParagraph(endParagraph))
            {
                endParagraph.Remove();
            }
        }

        private bool CanDeleteParagraph(Paragraph paragraph)
        {
            if (paragraph.Parent is TableCell)
            {
                // TableCell should have at least one paragraph element.
                var count = paragraph.Parent.ChildElements.Count(c => c is Paragraph);
                return count > 1;
            }

            return true;
        }

        internal virtual void Initialize()
        {
        }
    }
}