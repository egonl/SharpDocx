using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Models;

namespace SharpDocx
{
    public static class CodeBlockFactory
    {
        public static List<CodeBlock> Create(CharacterMap map, bool replaceCodeWithPlaceholder)
        {
            var codeBlocks = new List<CodeBlock>();

            var startTagIndex = map.Text.IndexOf("<%", 0);

            while (startTagIndex != -1)
            {
                var endTagIndex = map.Text.IndexOf("%>", startTagIndex + 2);
                if (endTagIndex == -1)
                {
                    throw new Exception("No end tag found for code.");
                }

                codeBlocks.Add(new CodeBlock
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
                for (var i = codeBlocks.Count - 1; i >= 0; --i)
                {
                    // Replace the code of each code block with an empty Text element.
                    var codeBlock = codeBlocks[i];
                    codeBlock.Placeholder = map.ReplaceWithText(codeBlock, null);
                    //codeBlock.Placeholder = map.ReplaceWithText(codeBlock, $"CB{i}");
                }
            }

            for (var i = 0; i < codeBlocks.Count; ++i)
            {
                var cb = codeBlocks[i];

                if (cb.Code.Replace(" ", "").StartsWith("if(") && codeBlocks[i].CurlyBracketLevelIncrement > 0)
                {
                    cb.Conditional = true;
                    cb.Condition = cb.GetExpressionInBrackets();

                    if (replaceCodeWithPlaceholder)
                    {
                        var bracketLevel = codeBlocks[i].CurlyBracketLevelIncrement;

                        for (var j = i + 1; j < codeBlocks.Count; ++j)
                        {
                            bracketLevel += codeBlocks[j].CurlyBracketLevelIncrement;

                            if (bracketLevel <= 0)
                            {
                                cb.EndConditionalPart = codeBlocks[j].Placeholder;
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

            // Update the map, since we modified the document.
            map.Recreate();

            return codeBlocks;
        }
    }
}