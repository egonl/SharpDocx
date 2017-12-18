using DocumentFormat.OpenXml;
using SharpDocx.Extensions;
using SharpDocx.Models;

namespace SharpDocx
{
    public class ElementAppender<T> where T : OpenXmlElement
    {
        private CodeBlock _elementCodeBlock;
        private T _lastElement;
        private T _templateElement;

        public void Append(CodeBlock currentCodeBlock)
        {
            if (_elementCodeBlock != currentCodeBlock)
            {
                // First paragraph/row only.
                _elementCodeBlock = currentCodeBlock;
                _templateElement = currentCodeBlock.Placeholder.GetParent<T>();
                var newParagraph = _templateElement.Clone() as T;
                _templateElement.InsertAfterSelf(newParagraph);
                _lastElement = newParagraph;
                _templateElement.Remove();
            }
            else
            {
                // Subsequent paragraphs/rows.
                var newParagraph = _templateElement.Clone() as T;
                _lastElement.InsertAfterSelf(newParagraph);
                _lastElement = newParagraph;
            }
        }
    }
}