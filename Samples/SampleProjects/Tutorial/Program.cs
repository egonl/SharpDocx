using System.IO;
using SharpDocx;

namespace Tutorial
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        private static void Main()
        {
            var viewPath = $"{BasePath}/Views/Tutorial.cs.docx";
            var documentPath = $"{BasePath}/Documents/Tutorial.docx";
            var imageDirectory = $"{BasePath}/Images";

#if DEBUG
            string documentViewer = null; // Try to find a document viewer automatically.
            //var documentViewer = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // Or specify a document viewer manually. 

            Ide.Start(viewPath, documentPath, null, null, f => f.ImageDirectory = imageDirectory, documentViewer);
#else
            DocumentBase document = DocumentFactory.Create(viewPath);
            document.ImageDirectory = imageDirectory;
            document.Generate(documentPath);
            Console.WriteLine($"Succesfully generated {documentPath} using view {viewPath}.");
#endif
        }
    }
}