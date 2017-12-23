using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDocx.CodeBlocks
{
    /// <summary>
    ///     ConditionalText is used to conditionally show or hide Word elements.
    /// </summary>
    public class ConditionalText : CodeBlock
    {
        public string Condition { get; set; }

        public Text EndConditionalPart { get; set; }

        public ConditionalText(string code) : base(code)
        {
        }
    }
}