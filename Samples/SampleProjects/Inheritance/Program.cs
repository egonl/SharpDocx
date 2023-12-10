using SharpDocx;
using System.IO;

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
            string documentViewer = null; // NET35 and NET45 will automatically search for a Docx viewer.
            //var documentViewer = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // NETCOREAPP3_1 and NET6_0 won't.

            Ide.Start(viewPath, documentPath, null, typeof(MyDocument), f => ((MyDocument)f).MyProperty = "The code", documentViewer);
#else

            var documentStreamPath = documentPath.Replace(".docx", ".stream.docx");
            MemoryStream ms = new MemoryStream(File.ReadAllBytes(viewPath));
            var myDocument = DocumentFactory.Create<MyDocument>((new FileInfo(viewPath)).Name, ms);
            myDocument.MyProperty = "The Code";

            // It's possible to generate a file or a stream.

            // 1. Generate a file
            // myDocument.Generate(documentPath);

            //2. Generate an output stream.
            using (var outputStream = myDocument.Generate(ms))
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