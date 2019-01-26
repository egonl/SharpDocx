using System;
using System.Diagnostics;
using System.IO;

namespace SharpDocx
{
    public static class Ide
    {
        public static void Start(
            string viewPath,
            string documentPath,
            object model = null,
            Type baseClassType = null,
            Action<DocumentBase> initializeDocument = null)
        {
            Console.WriteLine("Initializing SharpDocx IDE...");

            viewPath = Path.GetFullPath(viewPath);
            documentPath = Path.GetFullPath(documentPath);
            ConsoleKeyInfo keyInfo;

            do
            {
                try
                {
                    Console.WriteLine();
                    Console.WriteLine($"Compiling '{viewPath}'.");
                    var document = DocumentFactory.Create(viewPath, model, baseClassType, true);
                    initializeDocument?.Invoke(document);
                    document.Generate(documentPath);
                    Console.WriteLine($"Succesfully generated '{documentPath}'.");

                    try
                    {              
                        // Show the generated document.
                        Process.Start(documentPath);
                    }
                    catch
                    {
                        // Ignored.
                    }
                }
                catch (SharpDocxCompilationException e)
                {
                    Console.WriteLine(e.SourceCode);
                    Console.WriteLine(e.Errors);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }

                GC.Collect();
                Console.WriteLine("Press Esc to exit, any other key to retry . . .");
                keyInfo = Console.ReadKey(true);
            } while (keyInfo.Key != ConsoleKey.Escape);
        }
    }
}