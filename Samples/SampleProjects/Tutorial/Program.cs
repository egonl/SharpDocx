using SharpDocx;

namespace Tutorial
{
    internal class Program
    {
        static readonly string BasePath =  System.IO.Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"\..\..\..\..";

        private static void Main()
        {
            var viewPath = $"{BasePath}\\Views\\Tutorial.cs.docx";
            var documentPath = $"{BasePath}\\Documents\\Tutorial.docx";
            var imageDirectory = $"{BasePath}\\Images";

#if DEBUG
            Ide.Start(viewPath, documentPath, null, null, f => f.ImageDirectory = imageDirectory);
#else
            DocumentBase document = DocumentFactory.Create(viewPath);
            document.ImageDirectory = imageDirectory;
            document.Generate(documentPath);
#endif
        }
    }
}