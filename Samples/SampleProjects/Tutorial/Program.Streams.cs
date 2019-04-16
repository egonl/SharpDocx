using System;
using System.IO;
using SharpDocx;

namespace Tutorial
{
    internal class ProgramStream
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        private static void Main()
        {
            var viewPath = $"{BasePath}/Views/Tutorial.cs.docx";
            var documentPath = $"{BasePath}/Documents/Tutorial.Stream.docx";
            var imageDirectory = $"{BasePath}/Images";


            DocumentBase document = DocumentFactory.Create(viewPath);
            document.ImageDirectory = imageDirectory;
            var stream = document.Generate();

            using (var filestream = File.OpenWrite(documentPath))
            {
                var bytes = stream.ToArray();
                filestream.Write(bytes, 0, Convert.ToInt32(bytes.Length));
            }

        }
    }
}