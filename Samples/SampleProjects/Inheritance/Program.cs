using System.IO;
using SharpDocx;

namespace Inheritance
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        private static void Main()
        {
            var viewPath = $"{BasePath}/Views/Inheritance.cs.docx";
            var documentPath = $"{BasePath}/Documents/Inheritance.docx";

#if DEBUG
            Ide.Start(viewPath, documentPath, null, typeof(MyDocument), f => ((MyDocument) f).MyProperty = "The code");
#else
            var myDocument = DocumentFactory.Create<MyDocument>(viewPath);
            myDocument.MyProperty = "The Code";
            myDocument.Generate(documentPath);
#endif
        }
    }
}