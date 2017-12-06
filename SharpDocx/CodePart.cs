using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ActiveWord
{
    public class CodePart
    {
        public List<OpenXmlElement> Elements;

        public Text StartText;
        public int StartTagIndex;

        public Text EndText;
        public int EndTagIndex;

        public CodePart()
        {
            this.Elements = new List<OpenXmlElement>();
        }
    }
}