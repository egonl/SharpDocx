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

            //DocumentBase document = DocumentFactory.Create(viewPath);
            //document.Generate(documentPath);

            Ide.Start(viewPath, documentPath);
        }
    }
}