using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx;

using System.Collections.Generic;

namespace Inheritance
{
    public abstract class MyDocument : DocumentFileBase
    {
        public string MyProperty { get; set; }

        public new static List<string> GetUsingDirectives()
        {
            return new List<string>
            {
                "using Inheritance;"

                //"using static System.Math;"
                // Requires support for C# 6.
                // See https://stackoverflow.com/questions/31639602/using-c-sharp-6-features-with-codedomprovider-rosyln
            };
        }

        public new static List<string> GetReferencedAssemblies()
        {
            return new List<string>
            {
                typeof(MyDocument).Assembly.Location
            };
        }

        protected void CreateHyperlink(string text, string url)
        {
            // This method will be called from Inheritance.cs.docx.
            var id = $"r{Guid.NewGuid():N}";

            var hyperlink = new Hyperlink(
                new Run(
                    new RunProperties(new RunStyle { Val = "Hyperlink" }),
                    new Text(text)))
            {
                History = true,
                Id = id
            };

            Package.MainDocumentPart.AddHyperlinkRelationship(
                new Uri(url, UriKind.Absolute), true, id);

            CurrentCodeBlock.Placeholder.Parent.InsertAfterSelf(hyperlink);
        }
    }
}