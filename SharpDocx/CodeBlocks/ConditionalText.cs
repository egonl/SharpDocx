using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.CodeBlocks
{
    /// <summary>
    ///     ConditionalText is used to conditionally show or hide Word elements.
    /// </summary>
    internal class ConditionalText : CodeBlock
    {
        public string Condition { get; set; }

        public Text EndConditionalPart { get; set; }

        public ConditionalText(string code) : base(code)
        {
            Condition = code.GetExpressionInBrackets();
        }
    }
}