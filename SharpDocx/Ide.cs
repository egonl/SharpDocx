using System;

namespace SharpDocx
{
    public static class Ide
    {
        public static void Start(string viewPath, string documentPath, object model = null, Type baseClassType = null)
        {
            Console.WriteLine("Initializing SharpDocx IDE...");

            viewPath = System.IO.Path.GetFullPath(viewPath);
            documentPath = System.IO.Path.GetFullPath(documentPath);
            ConsoleKeyInfo keyInfo;

            do
            {
                try
                {
                    Console.WriteLine();
                    Console.WriteLine($"Compiling '{viewPath}'.");
                    var document = DocumentFactory.Create(viewPath, model, baseClassType, true);
                    document.Generate(documentPath);
                    Console.WriteLine($"Succesfully generated '{documentPath}'.");
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

                Console.WriteLine("Press Esc to exit, any other key to retry . . .");
                keyInfo = Console.ReadKey(true);
            } while (keyInfo.Key != ConsoleKey.Escape);
        }
    }
}