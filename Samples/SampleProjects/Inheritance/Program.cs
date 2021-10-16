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
            
            // It's possible to generate a file or a stream.
            
            // 1. Generate a file
            // myDocument.Generate(documentPath);

            //2. Generate an output stream.
            using (var outputStream = myDocument.Generate())
            {
                using (var outputFile = File.Open(documentPath, FileMode.Create))
                {
                    outputFile.Write(outputStream.GetBuffer(), 0, (int)outputStream.Length);
                }
            }
#endif
        }
    }
}