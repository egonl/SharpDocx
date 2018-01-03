using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Models;

namespace SharpDocx
{
    public class CharacterMap
    {
        private readonly List<OpenXmlElement> _elements = new List<OpenXmlElement>();

        private readonly List<Character> _map = new List<Character>();

        private readonly StringBuilder _textBuilder = new StringBuilder();

        private OpenXmlCompositeElement _rootElement;

        private string _text;

        public Character this[int index]
        {
            get
            {
                if (IsDirty)
                {
                    Recreate();
                }

                return _map[index];
            }
        }

        public string Text
        {
            get
            {
                if (IsDirty)
                {
                    Recreate();
                }

                return _text;
            }
        }

        private List<OpenXmlElement> Elements
        {
            get
            {
                if (IsDirty)
                {
                    Recreate();
                }

                return _elements;
            }
        }

        internal bool IsDirty { get; set; }

        public CharacterMap(OpenXmlCompositeElement ce)
        {
            _rootElement = ce;
            CreateMap(_rootElement);
            _text = _textBuilder.ToString();
            IsDirty = false;
        }

        private void Recreate()
        {
            _elements.Clear();
            _map.Clear();
            _textBuilder.Length = 0;
            CreateMap(_rootElement);
            _text = _textBuilder.ToString();
            IsDirty = false;
        }

        private void CreateMap(OpenXmlCompositeElement ce)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    CreateMap(child as OpenXmlCompositeElement);
                }
                else
                {
                    _elements.Add(child);
                }

                if (child is Paragraph || child is Break)
                {
                    _map.Add(new Character
                    {
                        Char = (char)10,
                        Element = child,
                        Index = -1
                    });

                    _textBuilder.Append((char)10);
                }

                if (child is Text)
                {
                    var t = child as Text;
                    for (var i = 0; i < t.Text.Length; ++i)
                    {
                        _map.Add(new Character
                        {
                            Char = t.Text[i],
                            Element = child,
                            Index = i
                        });
                    }

                    _textBuilder.Append(t.Text);
                }
            }

            _elements.Add(ce);
        }

        public int GetIndex(Text text)
        {
            // Can be used to get the index of a CodeBlock.Placeholder.
            // Then you can replace text that occurs after the code block only (instead of all text).
            var index = Elements.IndexOf(text);
            if (index == -1)
            {
                return -1;
            }

            for (var i = index; i >= 0; --i)
            {
                var t = Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return _map.IndexOf(_map.First(c => c.Element == t && c.Index == t.Text.Length - 1));
                }
            }

            for (var i = index + 1; i < Elements.Count; ++i)
            {
                var t = Elements[i] as Text;
                if (t != null && t.Text.Length > 0)
                {
                    return _map.IndexOf(_map.First(c => c.Element == t && c.Index == 0));
                }
            }

            return 0;
        }

        internal void Replace(string oldValue, string newValue, int startIndex = 0,
            StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            var i = Text.IndexOf(oldValue, startIndex, stringComparison);
            var dirty = i != -1;

            while (i != -1)
            {
                var part = new MapPart
                {
                    StartIndex = i,
                    EndIndex = i + oldValue.Length - 1
                };

                Replace(part, newValue);

                startIndex = i + newValue.Length;
                i = Text.IndexOf(oldValue, startIndex, stringComparison);
            }

            IsDirty = dirty;
        }

        private void Replace(MapPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            var parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (var i = Elements.IndexOf(endText); i >= Elements.IndexOf(startText); --i)
            {
                var element = Elements[i];

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

        internal Text ReplaceWithText(MapPart part, string newText)
        {
            var startText = this[part.StartIndex].Element as Text;
            var startIndex = this[part.StartIndex].Index;
            var endText = this[part.EndIndex].Element as Text;
            var endIndex = this[part.EndIndex].Index;

            var addedText = new Text
            {
                Text = newText,
                Space = SpaceProcessingModeValues.Preserve
            };

            var parents = new List<OpenXmlElement>();
            var parent = startText.Parent;
            while (parent != null)
            {
                parents.Add(parent);
                parent = parent.Parent;
            }

            for (var i = Elements.IndexOf(endText); i >= Elements.IndexOf(startText); --i)
            {
                var element = Elements[i];

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

        internal void Delete(OpenXmlElement startText, OpenXmlElement endText)
        {
            var parents = new List<OpenXmlElement>();
            AddParents(startText, parents);
            AddParents(endText, parents);

            for (var i = Elements.IndexOf(endText); i >= Elements.IndexOf(startText); --i)
            {
                var element = Elements[i];

                if (parents.Contains(element))
                {
                    // Do not remove parents.
                    continue;
                }

                element.Remove();
            }

            IsDirty = true;
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