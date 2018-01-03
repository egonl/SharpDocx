using System;
using System.Collections.Generic;
using SharpDocx.Extensions;

namespace SharpDocx.CodeBlocks
{
    internal class Directive : CodeBlock
    {
        public string Name { get; }

        public Dictionary<string, string> Attributes { get; } = new Dictionary<string, string>();

        public Directive(string code) : base(code)
        {
            var stringExpressions = new Dictionary<string, string>();
            string stringExpression;
            var startIndex = 0;

            do
            {
                stringExpression = code.GetExpressionInApostrophes(startIndex);
                if (stringExpression != null)
                {
                    // Replace attribute values enclosed in apostrophes with guids to avoid parsing problems later.
                    // E.g. attribute="The value of 1 + 1 = 2." becomes attribute=041fb3eedc5349d88437082a71f00ee5
                    var stringId = Guid.NewGuid().ToString("N");
                    stringExpressions.Add(stringId, stringExpression);
                    code = code.Replace($"\"{stringExpression}\"", stringId);
                    startIndex = code.IndexOf(stringId) + stringId.Length;
                }
            } while (stringExpression != null);

            var tagContents = code
                .Replace('\n', ' ')
                .RemoveRedundantWhitespace()
                .Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries);

            Name = tagContents[0].ToLower();

            for (var i = 1; i < tagContents.Length; ++i)
            {
                var attributeContents = tagContents[i].Split('=');

                Attributes.Add(
                    attributeContents[0].ToLower(),
                    stringExpressions.ContainsKey(attributeContents[1])
                        ? stringExpressions[attributeContents[1]]
                        : attributeContents[1]);
            }
        }
    }
}