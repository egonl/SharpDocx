using SharpDocx;

namespace Tutorial
{
    internal class Program
    {
        static readonly string BasePath =  System.IO.Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"\..\..\..\..";

        private static void Main(string[] args)
        {
            var viewPath = $"{BasePath}\\Views\\Tutorial.cs.docx";
            var documentPath = $"{BasePath}\\Documents\\Tutorial.docx";

#if DEBUG
            Ide.Start(viewPath, documentPath);
#else
            DocumentBase document = DocumentFactory.Create(viewPath);
            document.Generate(documentPath);
#endif
        }
    }
}