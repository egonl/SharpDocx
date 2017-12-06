using SharpDocx.Extensions;
using SharpDocx.Models;
using DocumentFormat.OpenXml;

namespace SharpDocx
{
    public class ElementAppender<T> where T : OpenXmlElement
    {
        private T lastElement;
        private CodeBlock elementCodeBlock;
        private T templateElement;

        public void Append(CodeBlock currentCodeBlock)
        {
            if (this.elementCodeBlock != currentCodeBlock)
            {
                // First paragraph/row only.
                this.elementCodeBlock = currentCodeBlock;
                this.templateElement = currentCodeBlock.Placeholder.GetParent<T>();
                var newParagraph = this.templateElement.Clone() as T;
                this.templateElement.InsertAfterSelf(newParagraph);
                this.lastElement = newParagraph;
                this.templateElement.Remove();
            }
            else
            {
                // Subsequent paragraphs/rows.
                var newParagraph = this.templateElement.Clone() as T;
                this.lastElement.InsertAfterSelf(newParagraph);
                this.lastElement = newParagraph;
            }
        }
    }
}