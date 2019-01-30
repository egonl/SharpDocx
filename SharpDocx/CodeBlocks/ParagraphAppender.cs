using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.CodeBlocks
{
    public class ParagraphAppender : CodeBlock
    {
        private Appender  _appender;
        private readonly Body _body = new Body();

        public ParagraphAppender(string code) : base(code)
        {
            CurrentInsertionPoint = new InsertionPoint();
        }

        internal override void Initialize()
        {
            base.Initialize();

            var previousElement = Placeholder.GetParent<Paragraph>().PreviousSibling() as OpenXmlCompositeElement;
            previousElement.SetAttribute(new OpenXmlAttribute { LocalName = "IpId", Value = CurrentInsertionPoint.Id });
            CurrentInsertionPoint.Element = previousElement;

            var p = Placeholder.GetParent<Paragraph>();
            p.Remove();
            _body.InsertAt(p, 0);
            _appender = new Appender(_body);
        }

        public void Append()
        {
            CurrentInsertionPoint.Element = _appender.Append(CurrentInsertionPoint.Element);
        }
    }
}
