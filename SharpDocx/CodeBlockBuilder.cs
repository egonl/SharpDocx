using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using SharpDocx.Models;

namespace SharpDocx
{
    internal class CodeBlockBuilder
    {
        public CharacterMap BodyMap { get; }

        public List<CodeBlock> CodeBlocks { get; }

        public CodeBlockBuilder(WordprocessingDocument package, bool replaceCodeWithPlaceholder)
        {
            this.BodyMap = CharacterMap.Create(package.MainDocumentPart.Document.Body);
            this.CodeBlocks = CodeBlockFactory.Create(this.BodyMap, replaceCodeWithPlaceholder);
            if (replaceCodeWithPlaceholder)
            {
                // Update the map, since we modified the document.
                this.BodyMap.Recreate();
            }

            foreach (var footerPart in package.MainDocumentPart.FooterParts)
            {
                var footerMap = CharacterMap.Create(footerPart.Footer);
                var footerBlocks = CodeBlockFactory.Create(footerMap, replaceCodeWithPlaceholder);
                this.CodeBlocks.InsertRange(0, footerBlocks);
            }

            foreach (var headerPart in package.MainDocumentPart.HeaderParts)
            {
                var headerMap = CharacterMap.Create(headerPart.Header);
                var headerBlocks = CodeBlockFactory.Create(headerMap, true);
                this.CodeBlocks.InsertRange(0, headerBlocks);
            }
        }
    }
}
