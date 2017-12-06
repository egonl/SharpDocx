using DocumentFormat.OpenXml;

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
    }
}
