using System.IO;
using SharpDocx;

namespace Inheritance
{
    internal class Program
    {
        private static readonly string BasePath = Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"\..\..\..\..";

        private static void Main(string[] args)
        {
            var viewPath = $"{BasePath}\\Views\\Inheritance.cs.docx";
            var documentPath = $"{BasePath}\\Documents\\Inheritance.docx";

            //Ide.Start(viewPath, documentPath, null, typeof(MyDocument));

            var myDocument = DocumentFactory.Create<MyDocument>(viewPath);
            myDocument.MyProperty = "excellent";
            myDocument.Generate(documentPath);
        }
    }
}