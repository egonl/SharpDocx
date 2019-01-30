using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;
using System.Collections.Generic;
using System.Linq;

namespace SharpDocx.CodeBlocks
{
    /// <summary>
    /// A TextBlock is a CodeBlock with an opening curly brace and @, textual content and a CodeBlock with a closing curly brace. E.g.:
    /// <code>
    /// &lt;% for (int i = 0; i &lt; 10; ++i) {@ %&gt;
    ///   Textual content.
    /// &lt;% } %&gt;
    /// </code>
    /// </summary>
    public class TextBlock : CodeBlock
    {
        internal CodeBlock EndingCodeBlock;

        private Appender _appender;
        private readonly Body _body = new Body();

        public TextBlock(string code) : base(code)
        {
            CurrentInsertionPoint = new InsertionPoint();
        }

        internal override void Initialize()
        {
            base.Initialize();

            var previousElement = Placeholder.GetElementBlockLevelParent().PreviousSibling() as OpenXmlCompositeElement;
            previousElement.SetAttribute(new OpenXmlAttribute { LocalName = "IpId", Value = CurrentInsertionPoint.Id });
            CurrentInsertionPoint.Element = previousElement;

            GetBody(StartText, EndingCodeBlock.Placeholder);
            _appender = new Appender(_body);
        }

        public void Append(List<CodeBlock> codeBlocks)
        {
            CurrentInsertionPoint.Element = _appender.Append(CurrentInsertionPoint.Element);

            // Update insertion points of *other* code blocks within this text block (e.g. the insert point of a row appender).
            var childInsertionPoints = _appender.GetInsertionPoints();

            foreach (var childInsertionPoint in childInsertionPoints)
            {
                codeBlocks.First(x => x.CurrentInsertionPoint?.Id == childInsertionPoint.Id)
                    .CurrentInsertionPoint.Element = childInsertionPoint.Element;
            }
        }

        internal void GetBody(OpenXmlElement startText, OpenXmlElement endText)
        {
            var startParent = startText.GetElementBlockLevelParent();
            var endParent = endText.GetElementBlockLevelParent();

            if (startParent == endParent)
            {
                startParent.Remove();
                _body.InsertAt(startParent, 0);
                return;
            }

            var elementsToMove = new List<OpenXmlCompositeElement>();
            var nextElement = startParent;
            while (nextElement != endParent)
            {
                elementsToMove.Add(nextElement);
                nextElement = nextElement.NextSibling() as OpenXmlCompositeElement;
            }
            elementsToMove.Add(endParent);

            foreach (var element in elementsToMove)
            {
                element.Remove();
                _body.InsertAt(element, _body.ChildElements.Count);
            }
        }
    }
}