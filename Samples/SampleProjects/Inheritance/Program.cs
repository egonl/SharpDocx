using System.IO;
using SharpDocx;
#if NET35
using SharpDocx.Extensions;
#endif

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

            Ide.Start(viewPath, documentPath, null, typeof(MyDocument), f => ((MyDocument) f).MyProperty = "The code", documentViewer);
#else
            // 1 - It's possible to use a file or a stream for a view.

            // 1a - Use a file for a view
            //var myDocument = DocumentFactory.Create<MyDocument>(viewPath);

            // 1b - Or use a stream for a view
            using var viewStream = File.OpenRead(viewPath);
            var myDocument = DocumentFactory.Create<MyDocument>(viewStream);

            myDocument.MyProperty = "The Code";

            // 2 - It's also possible to generate a file or a stream.

            // 2a - Generate a file
            //myDocument.Generate(documentPath);

            // 2b - Or generate an output stream.
            using var outputStream = myDocument.Generate();
            using var outputFile = File.Open(documentPath, FileMode.Create);
            outputStream.CopyTo(outputFile);
#endif
        }
    }
}