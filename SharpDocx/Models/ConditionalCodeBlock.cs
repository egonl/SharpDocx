using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDocx.Models
{
    /// <summary>
    /// ConditionalCodeBlock is used to conditionally show or hide Word elements.
    /// </summary>
    public class ConditionalCodeBlock : CodeBlock
    {
        public string Condition { get; set; }

        public Text EndConditionalPart { get; set; }

        public ConditionalCodeBlock(string code) : base(code)
        {            
        }
    }
}