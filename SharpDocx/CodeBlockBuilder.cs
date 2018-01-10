using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.CodeBlocks;
using SharpDocx.Extensions;
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
                var headerMap = new CharacterMap(headerPart.Header);
                AppendCodeBlocks(headerMap, replaceCodeWithPlaceholder);
            }

            foreach (var footerPart in package.MainDocumentPart.FooterParts)
            {
                var footerMap = new CharacterMap(footerPart.Footer);
                AppendCodeBlocks(footerMap, replaceCodeWithPlaceholder);
            }

            if (package.MainDocumentPart.EndnotesPart != null)
            {
                var endnotesMap = new CharacterMap(package.MainDocumentPart.EndnotesPart.Endnotes);
                AppendCodeBlocks(endnotesMap, replaceCodeWithPlaceholder);
            }

            if (package.MainDocumentPart.FootnotesPart != null)
            {
                var footnotesMap = new CharacterMap(package.MainDocumentPart.FootnotesPart.Footnotes);
                AppendCodeBlocks(footnotesMap, replaceCodeWithPlaceholder);
            }

            BodyMap = new CharacterMap(package.MainDocumentPart.Document.Body);
            AppendCodeBlocks(BodyMap, replaceCodeWithPlaceholder);

            if (replaceCodeWithPlaceholder)
            {
                // Set the dirty flag on the map, since the document is modified.
                BodyMap.IsDirty = true;
            }
        }

        private void AppendCodeBlocks(CharacterMap map, bool replaceCodeWithPlaceholder)
        {
            var mapParts = new Dictionary<CodeBlock, MapPart>();

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
                cb.StartText = map[startTagIndex].Element as Text;
                cb.EndText = map[endTagIndex + 1].Element as Text;
                CodeBlocks.Add(cb);
                mapParts.Add(cb, new MapPart { StartIndex = startTagIndex, EndIndex = endTagIndex + 1});
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
                cb.Placeholder = map.ReplaceWithText(mapParts[cb], null);
                //cb.Placeholder = map.ReplaceWithText(mapParts[cb], $"CB{i}");
            }

            for (var i = firstCodeBlockIndex; i < CodeBlocks.Count; ++i)
            {
                // Find out where conditional content ends.
                var cb = CodeBlocks[i] as ConditionalText;
                if (cb != null)
                {
                    var bracketLevel = cb.Code.GetCurlyBracketLevelIncrement();

                    for (var j = i + 1; j < CodeBlocks.Count; ++j)
                    {
                        bracketLevel += CodeBlocks[j].Code.GetCurlyBracketLevelIncrement();

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

            for (var i = firstCodeBlockIndex; i < CodeBlocks.Count; ++i)
            {
                var appender = CodeBlocks[i] as Appender;
                appender?.Initialize();
            }
        }

        private static CodeBlock GetCodeBlock(string code)
        {
            CodeBlock cb;

            code = code.RemoveRedundantWhitespace();

            if (code.StartsWith("@"))
            {
                cb = new Directive(code.Substring(1));
            }
            else if (code.Contains("AppendParagraph") || code.Contains("AppendRow"))
            {
                // TODO: match whole words only. 
                cb = new Appender(code);
            }
            else if (code.Replace(" ", String.Empty).StartsWith("if(") &&
                     code.GetCurlyBracketLevelIncrement() > 0)
            {
                cb = new ConditionalText(code);
            }
            else
            {
                cb = new CodeBlock(code);
            }

            return cb;
        }
    }
}