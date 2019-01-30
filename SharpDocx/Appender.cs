using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;

namespace SharpDocx
{
    internal class Appender
    {
        private readonly List<Text> _placeholders;
        private readonly Body _templateBody;
        private Body _workingBody;

        public Appender(Body templateBody)
        {
            _workingBody = templateBody;
            _placeholders = GetPlaceholders(_workingBody);
            _templateBody = _workingBody.Clone() as Body;
        }

        internal OpenXmlCompositeElement Append(OpenXmlCompositeElement insertionPoint)
        {
            var elementsToMove = new List<OpenXmlCompositeElement>();
            foreach (var element in _workingBody.ChildElements)
            {
                elementsToMove.Add(element as OpenXmlCompositeElement);
            }
            
            foreach (var element in elementsToMove)
            {
                element.Remove();

                if (element.HasText())
                {
                    // Ignore empty paragraphs.
                    insertionPoint = insertionPoint.InsertAfterSelf(element);
                }
            }

            CreateNewWorkingBody();
            return insertionPoint;
        }

        internal List<InsertionPoint> GetInsertionPoints()
        {
            return InsertionPoint.FindAll(_workingBody);
        }

        // Return a list of possible placeholders (i.e. empty text elements).
        private static List<Text> GetPlaceholders(Body body)
        {
            var list = new List<Text>();
            foreach (var element in body.ChildElements)
            {
                GetPlaceholdersRecursive(element as OpenXmlCompositeElement, list);
            }
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
                    // Empty Text elements are assumed to be Placeholders.
                    placeholders.Add(text);
                }
            }
        }

        private void CreateNewWorkingBody()
        {
            _workingBody = _templateBody.Clone() as Body;
            
            // The working body should *always* contain the placeholders.
            // So copy the placeholders to the new working body.
            var clonedPlaceholders = GetPlaceholders(_workingBody);

            for (var i = 0; i < clonedPlaceholders.Count; ++i)
            {
                // Make a copy of the placeholder and insert this copy in the document.
                _placeholders[i].InsertAfterSelf(_placeholders[i].Clone() as Text);
                
                // Remove the original placeholder from the document and reset the it.
                _placeholders[i].Remove();
                _placeholders[i].Text = null;

                // Add the placeholder to the new working body.
                clonedPlaceholders[i].InsertAfterSelf(_placeholders[i]);

                // Remove the cloned placeholder from the working body.
                clonedPlaceholders[i].Remove();
            }
        }
    }
}