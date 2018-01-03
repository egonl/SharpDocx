using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx.CodeBlocks
{
    internal class Appender : CodeBlock
    {
        private OpenXmlElement _lastElement;
        private OpenXmlCompositeElement _newElement;
        private List<Text> _placeholders = new List<Text>();
        private OpenXmlCompositeElement _templateElement;

        public Appender(string code) : base(code)
        {
        }

        // Get the placeholders in the template and store a pristine copy of the template. 
        internal void Initialize()
        {
            OpenXmlCompositeElement templateElement = null;

            if (Code.Contains("AppendParagraph"))
            {
                templateElement = Placeholder.GetParent<Paragraph>();
            }
            else if (Code.Contains("AppendRow"))
            {
                templateElement = Placeholder.GetParent<TableRow>();
            }

            _newElement = templateElement;
            _placeholders = GetPlaceholders(templateElement);
            _templateElement = templateElement.Clone() as OpenXmlCompositeElement;
        }

        internal void Append<T>() where T : OpenXmlCompositeElement
        {
            // Clone _newElement, because we're going to use the placeholders currently in it again.
            var clonedNewElement = _newElement.Clone() as T;

            if (_lastElement == null)
            {
                // First call to Append: _newElement points to the original template.
                _lastElement = _newElement.InsertAfterSelf(clonedNewElement);

                // Remove the original template from the document.
                _newElement.Remove();
            }
            else
            {
                // Subsequent calls to Append.
                _lastElement = _lastElement.InsertAfterSelf(clonedNewElement);
            }

            // Create a pristine new element and initialize it with the placeholders. 
            _newElement = _templateElement.Clone() as T;
            SetPlaceholders(_newElement);
        }

        // Return a list of possible placeholders (i.e. empty text elements).
        private static List<Text> GetPlaceholders(OpenXmlCompositeElement ce)
        {
            var list = new List<Text>();
            GetPlaceholdersRecursive(ce, list);
            return list;
        }

        private static void GetPlaceholdersRecursive(OpenXmlCompositeElement ce, List<Text> placeholders)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    GetPlaceholdersRecursive(child as OpenXmlCompositeElement, placeholders);
                }
                var text = child as Text;
                if (text?.Text.Length == 0)
                {
                    placeholders.Add(text);
                }
            }
        }

        // Replace the empty text elements in newElement with empty placeholders. 
        private void SetPlaceholders(OpenXmlCompositeElement newElement)
        {
            var clonedPlaceholders = GetPlaceholders(newElement);
            for (var i = 0; i < clonedPlaceholders.Count; ++i)
            {
                _placeholders[i].Remove();
                _placeholders[i].Text = null;
                clonedPlaceholders[i].InsertAfterSelf(_placeholders[i]);
                clonedPlaceholders[i].Remove();
            }
        }
    }
}