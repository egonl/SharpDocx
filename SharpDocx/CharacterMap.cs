using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharpDocx.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDocx
{
    public class CharacterMap
    {
        public List<OpenXmlElement> Elements = new List<OpenXmlElement>();

        private readonly List<Character> map = new List<Character>();

        public string Text;

        private readonly StringBuilder textBuilder = new StringBuilder();

        private OpenXmlCompositeElement rootElement;

        public Character this[int index] => this.map[index];

        public static CharacterMap Create(OpenXmlCompositeElement ce)
        {
            var m = new CharacterMap {rootElement = ce};
            m.CreateMap(m.rootElement);
            m.Text = m.textBuilder.ToString();
            return m;
        }

        public void Recreate()
        {
            this.Elements.Clear();
            this.map.Clear();
            this.textBuilder.Length = 0;
            CreateMap(this.rootElement);
            this.Text = this.textBuilder.ToString();
        }

        private void CreateMap(OpenXmlCompositeElement ce)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    CreateMap(child as OpenXmlCompositeElement);
                    if (child is Paragraph)
                    {
                        this.map.Add(new Character
                        {
                            Char = (char) 10,
                            Element = child,
                            Index = -1
                        });

                        this.textBuilder.Append((char) 10);
                    }
                }
                else
                {
                    this.Elements.Add(child);
                }

                if (child is Text)
                {
                    var t = child as Text;
                    for (var i = 0; i < t.Text.Length; ++i)
                    {
                        this.map.Add(new Character
                        {
                            Char = t.Text[i],
                            Element = child,
                            Index = i
                        });
                    }

                    this.textBuilder.Append(t.Text);
                }
            }

            this.Elements.Add(ce);
        }

        public int GetIndex(Text text)
        {
            var index = this.Elements.IndexOf(text);
            if (index == -1)
            {
                return -1;
            }

            for (int i = index; i >= 0; --i)
            {
                var t = this.Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return this.map.IndexOf(this.map.First(c => c.Element == t && c.Index == t.Text.Length - 1));
                }
            }

            for (int i = index + 1; i < this.Elements.Count; ++i)
            {
                var t = this.Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return this.map.IndexOf(this.map.First(c => c.Element == t && c.Index == 0));
                }
            }

            return 0;
        }

        public void Replace(string oldValue, string newValue, int startIndex = 0, StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            var i = this.Text.IndexOf(oldValue, startIndex, stringComparison);
            while (i != -1)
            {
                TextPart part = new TextPart
                {
                    StartIndex = i,
                    EndIndex = i + oldValue.Length - 1,
                };

                Replace(part, newValue);

                startIndex = i + newValue.Length;
                i = this.Text.IndexOf(oldValue, startIndex, stringComparison);
            }

            this.Recreate();
        }

        private void Replace(TextPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            List<OpenXmlElement> parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (int i = this.Elements.IndexOf(endText); i >= this.Elements.IndexOf(startText); --i)
            {
                var element = this.Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }

                if (element == startText)
                {
                    startText.Space = SpaceProcessingModeValues.Preserve;

                    if (startText == endText)
                    {
                        string postScriptum = null;
                        if (endIndex + 1 != startText.Text.Length)
                        {
                            postScriptum = startText.Text.Substring(endIndex + 1);
                        }

                        startText.Text = startText.Text.Substring(0, startIndex);
                        if (!string.IsNullOrEmpty(newText))
                        {
                            startText.Text += newText;
                        }
                        startText.Text += postScriptum;
                    }
                    else
                    {
                        startText.Text = startText.Text.Substring(0, startIndex);
                        if (!string.IsNullOrEmpty(newText))
                        {
                            startText.Text += newText;
                        }
                    }
                }
                else if (element == endText)
                {
                    endText.Space = SpaceProcessingModeValues.Preserve;
                    endText.Text = endText.Text.Substring(endIndex + 1);
                }
                else
                {
                    element.Remove();
                }
            }
        }

        public Text ReplaceWithText(TextPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            Text addedText = new Text
            {
                Text = newText,
                Space = SpaceProcessingModeValues.Preserve,
            };

            List<OpenXmlElement> parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (int i = this.Elements.IndexOf(endText); i >= this.Elements.IndexOf(startText); --i)
            {
                var element = this.Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }

                if (element == startText)
                {
                    startText.Space = SpaceProcessingModeValues.Preserve;

                    if (startText == endText)
                    {
                        string postScriptum = null;
                        if (endIndex + 1 != startText.Text.Length)
                        {
                            postScriptum = startText.Text.Substring(endIndex + 1);
                        }

                        startText.Text = startText.Text.Substring(0, startIndex);

                        if (!string.IsNullOrEmpty(postScriptum))
                        {
                            startText.InsertAfterSelf(new Text
                            {
                                Text = postScriptum,
                                Space = SpaceProcessingModeValues.Preserve
                            });
                        }

                        startText.InsertAfterSelf(addedText);
                    }
                    else
                    {
                        startText.Text = startText.Text.Substring(0, startIndex);
                        startText.InsertAfterSelf(addedText);
                    }
                }
                else if (element == endText)
                {
                    endText.Space = SpaceProcessingModeValues.Preserve;
                    endText.Text = endText.Text.Substring(endIndex + 1);
                }
                else
                {
                    element.Remove();
                }
            }

            return addedText;
        }

        public void Delete(OpenXmlElement startText, OpenXmlElement endText)
        {
            List<OpenXmlElement> parents = new List<OpenXmlElement>();
            AddParents(startText, parents);
            AddParents(endText, parents);

            for (int i = this.Elements.IndexOf(endText); i >= this.Elements.IndexOf(startText); --i)
            {
                var element = this.Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }

                element.Remove();
            }
        }

        private static void AddParents(OpenXmlElement element, List<OpenXmlElement> parents)
        {
            var parent = element;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }
        }
    }
}
