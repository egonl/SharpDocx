using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Models;

namespace SharpDocx
{
    internal class CodeBlockBuilder
    {
        public CharacterMap BodyMap { get; }

        public List<CodeBlock> CodeBlocks { get; }

        public CodeBlockBuilder(WordprocessingDocument package, bool replaceCodeWithPlaceholder)
        {
            this.CodeBlocks = new List<CodeBlock>();

            foreach (var headerPart in package.MainDocumentPart.HeaderParts)
            {
                var headerMap = CharacterMap.Create(headerPart.Header);
                AppendCodeBlocks(headerMap, replaceCodeWithPlaceholder);
            }

            foreach (var footerPart in package.MainDocumentPart.FooterParts)
            {
                var footerMap = CharacterMap.Create(footerPart.Footer);
                AppendCodeBlocks(footerMap, replaceCodeWithPlaceholder);
            }

            this.BodyMap = CharacterMap.Create(package.MainDocumentPart.Document.Body);
            AppendCodeBlocks(this.BodyMap, replaceCodeWithPlaceholder);

            if (replaceCodeWithPlaceholder)
            {
                // Set the dirty flag on the map, since the document is modified.
                this.BodyMap.IsDirty = true;
            }
        }

        private void AppendCodeBlocks(CharacterMap map, bool replaceCodeWithPlaceholder)
        {
            var startTagIndex = map.Text.IndexOf("<%", 0);
            var firstCodeBlockIndex = this.CodeBlocks.Count;

            while (startTagIndex != -1)
            {
                var endTagIndex = map.Text.IndexOf("%>", startTagIndex + 2);
                if (endTagIndex == -1)
                {
                    throw new Exception("No end tag found for code.");
                }

                this.CodeBlocks.Add(new CodeBlock
                {
                    StartIndex = startTagIndex,
                    StartText = map[startTagIndex].Element as Text,
                    EndIndex = endTagIndex + 1,
                    EndText = map[endTagIndex + 1].Element as Text,
                    Code = map.Text.Substring(startTagIndex + 2, endTagIndex - startTagIndex - 2)
                });

                startTagIndex = map.Text.IndexOf("<%", endTagIndex + 2);
            }

            if (replaceCodeWithPlaceholder)
            {
                for (var i = this.CodeBlocks.Count - 1; i >= firstCodeBlockIndex; --i)
                {
                    // Replace the code of each code block with an empty Text element.
                    var codeBlock = this.CodeBlocks[i];
                    codeBlock.Placeholder = map.ReplaceWithText(codeBlock, null);
                    //codeBlock.Placeholder = map.ReplaceWithText(codeBlock, $"CB{i}");
                }
            }

            for (var i = firstCodeBlockIndex; i < this.CodeBlocks.Count; ++i)
            {
                var cb = this.CodeBlocks[i];

                if (cb.Code.Replace(" ", "").StartsWith("if(") && this.CodeBlocks[i].CurlyBracketLevelIncrement > 0)
                {
                    cb.Conditional = true;
                    cb.Condition = cb.GetExpressionInBrackets();

                    if (replaceCodeWithPlaceholder)
                    {
                        var bracketLevel = this.CodeBlocks[i].CurlyBracketLevelIncrement;

                        for (var j = i + 1; j < this.CodeBlocks.Count; ++j)
                        {
                            bracketLevel += this.CodeBlocks[j].CurlyBracketLevelIncrement;

                            if (bracketLevel <= 0)
                            {
                                cb.EndConditionalPart = this.CodeBlocks[j].Placeholder;
                                break;
                            }
                        }

                        if (cb.EndConditionalPart == null)
                        {
                            throw new Exception("Conditional block is not terminated with '<% } %>'.");
                        }
                    }
                }
            }
        }
    }
}
