﻿using System;
using System.IO;
using SharpDocx;
#if NET35 || NET40
using System.Net;
#endif

namespace Tutorial
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        private static void Main()
        {
#if NET35 || NET40
            // The Tutorial fetches an image from https://github.com/egonl/SharpDocx/blob/master/Samples/Images/test3.emf?raw=true,
            // which requires TLS 1.2. 
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;
#endif
            var viewPath = $"{BasePath}/Views/Tutorial.cs.docx";
            var documentPath = $"{BasePath}/Documents/Tutorial.docx";
            var imageDirectory = $"{BasePath}/Images";

#if DEBUG
            string documentViewer = null; // .NET Framework 3.5 or greater will automatically search for a Docx viewer.
            //var documentViewer = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // NETCOREAPP3_1 and NET6_0 won't.

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