using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpDocx
{
    public class InsertionPoint
    {
        public string Id = Guid.NewGuid().ToString("N");
        public OpenXmlCompositeElement Element;

        public static List<InsertionPoint> FindAll(OpenXmlCompositeElement elements)
        {
            var list = new List<InsertionPoint>();
            FindAll(elements, list);
            return list;
        }

        private static void FindAll(OpenXmlCompositeElement elements, List<InsertionPoint> list)
        {
            foreach (var element in elements.ChildElements)
            {
                var insertionPointId = element.ExtendedAttributes.Where(x => x.LocalName == "IpId").Select(x => x.Value).FirstOrDefault();
                if (!string.IsNullOrEmpty(insertionPointId))
                {
                    list.Add(new InsertionPoint { Element = element as OpenXmlCompositeElement, Id = insertionPointId });
                }

                if (element.HasChildren)
                {
                    FindAll(element as OpenXmlCompositeElement, list);
                }
            }
        }
    }
}
