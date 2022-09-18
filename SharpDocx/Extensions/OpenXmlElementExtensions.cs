using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SharpDocx.Extensions
{
    public static class OpenXmlElementExtensions
    {
        public static T GetParent<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            var e = element.Parent;

            while (e != null)
            {
                if (e is T t)
                {
                    return t;
                }

                e = e.Parent;
            }

            return null;
        }

        public static List<T> GetAllElements<T>(this OpenXmlElement e) where T : OpenXmlElement
        {
            var list = new List<T>();
            GetAllElementsRecursive(e, list);
            return list;
        }

        private static void GetAllElementsRecursive<T>(OpenXmlElement e, List<T> list) where T : OpenXmlElement
        {
            foreach (var child in e.ChildElements)
            {
                if (child.HasChildren)
                {
                    GetAllElementsRecursive(child as OpenXmlElement, list);
                }
                else if (child is T)
                {
                    list.Add(child as T);
                }
            }
        }


        public static string GetText(this OpenXmlCompositeElement ce)
        {
            var sb = new StringBuilder();
            GetTextRecursive(ce, sb);
            return sb.ToString();
        }

        private static void GetTextRecursive(OpenXmlCompositeElement ce, StringBuilder sb)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    GetTextRecursive(child as OpenXmlCompositeElement, sb);
                }
                else if (child is Text t)
                {
                    sb.Append(t.Text);
                }
            }
        }

        public static bool HasText(this OpenXmlCompositeElement ce)
        {
            foreach (var child in ce.ChildElements)
            {
                if (child.HasChildren)
                {
                    var hasText = HasText(child as OpenXmlCompositeElement);
                    if (hasText)
                    {
                        return true;
                    }
                }
                else if (child is Text t && t.Text.Length > 0)
                {
                    return true;
                }
                else if (child is Break b && b.Type != null && b.Type == BreakValues.Page)
                {
                    // Only preserve page breaks, see issue #36.
                    return true;
                }
            }

            return false;
        }

        internal static OpenXmlCompositeElement GetElementBlockLevelParent(this OpenXmlElement element)
        {
            // Block level elements: see https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.body?view=openxml-2.8.1
            var parent = element.GetParent<Paragraph>() as OpenXmlCompositeElement;
            if (parent != null)
            {
                return parent;
            }

            parent = element.GetParent<Table>() as OpenXmlCompositeElement;
            if (parent != null)
            {
                return parent;
            }

            return null;
        }
    }
}