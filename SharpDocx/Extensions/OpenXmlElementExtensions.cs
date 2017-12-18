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
                if (e is T)
                {
                    return e as T;
                }

                e = e.Parent;
            }

            return null;
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
                else if (child is Text)
                {
                    sb.Append((child as Text).Text);
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
                else if (child is Text && (child as Text).Text.Length > 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}