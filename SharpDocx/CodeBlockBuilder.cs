using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Models;

namespace SharpDocx
{
    internal class CodeBlockBuilder
    {
        public List<CodeBlock> CodeBlocks { get; }

        public CharacterMap BodyMap { get; }

        public CodeBlockBuilder(WordprocessingDocument package, bool replaceCodeWithPlaceholder)
        {
            CodeBlocks = new List<CodeBlock>();

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

            BodyMap = CharacterMap.Create(package.MainDocumentPart.Document.Body);
            AppendCodeBlocks(BodyMap, replaceCodeWithPlaceholder);

            if (replaceCodeWithPlaceholder)
            {
                // Set the dirty flag on the map, since the document is modified.
                BodyMap.IsDirty = true;
            }
        }

        private void AppendCodeBlocks(CharacterMap map, bool replaceCodeWithPlaceholder)
        {
            var startTagIndex = map.Text.IndexOf("<%", 0);
            var firstCodeBlockIndex = CodeBlocks.Count;

            while (startTagIndex != -1)
            {
                var endTagIndex = map.Text.IndexOf("%>", startTagIndex + 2);
                if (endTagIndex == -1)
                {
                    throw new Exception("No end tag found for code.");
                }

                var code = map.Text.Substring(startTagIndex + 2, endTagIndex - startTagIndex - 2);
                var cb = GetCodeBlock(code);
                cb.StartIndex = startTagIndex;
                cb.StartText = map[startTagIndex].Element as Text;
                cb.EndIndex = endTagIndex + 1;
                cb.EndText = map[endTagIndex + 1].Element as Text;
                cb.Code = code;
                CodeBlocks.Add(cb);

                startTagIndex = map.Text.IndexOf("<%", endTagIndex + 2);
            }

            if (!replaceCodeWithPlaceholder)
            {
                return;
            }

            for (var i = CodeBlocks.Count - 1; i >= firstCodeBlockIndex; --i)
            {
                // Replace the code of each code block with an empty Text element.
                var cb = CodeBlocks[i];
                cb.Placeholder = map.ReplaceWithText(cb, null);
                //cb.Placeholder = map.ReplaceWithText(cb, $"CB{i}");
            }

            for (var i = firstCodeBlockIndex; i < CodeBlocks.Count; ++i)
            {
                // Find out where conditional content ends.
                var cb = CodeBlocks[i] as ConditionalCodeBlock;
                if (cb != null)
                {
                    var bracketLevel = cb.CurlyBracketLevelIncrement;

                    for (var j = i + 1; j < CodeBlocks.Count; ++j)
                    {
                        bracketLevel += CodeBlocks[j].CurlyBracketLevelIncrement;

                        if (bracketLevel <= 0)
                        {
                            cb.EndConditionalPart = CodeBlocks[j].Placeholder;
                            break;
                        }
                    }

                    if (cb.EndConditionalPart == null)
                    {
                        throw new Exception("Conditional code block is not terminated with '<% } %>'.");
                    }
                }
            }
        }

        private static CodeBlock GetCodeBlock(string code)
        {
            CodeBlock cb;

            if (code.Replace(" ", "").StartsWith("if(") && CodeBlock.GetCurlyBracketLevelIncrement(code) > 0)
            {
                cb = new ConditionalCodeBlock
                {
                    Condition = CodeBlock.GetExpressionInBrackets(code),
                };
            }
            else
            {
                cb = new CodeBlock();
            }

            return cb;
        }
    }
}