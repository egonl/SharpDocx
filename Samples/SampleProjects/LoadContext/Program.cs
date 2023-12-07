using SharpDocx;
using System;
using System.IO;
using System.Linq;
using System.Runtime.Loader;

namespace LoadContext
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        public static void Main()
        {
            ExecuteTemplates(out var loadContextRef);

            while (loadContextRef.IsAlive)
            {
                Console.WriteLine("Reference is still alive.");
                //Console.WriteLine($"Loaded AssemblyCount: {AppDomain.CurrentDomain.GetAssemblies().Length}");
                //Thread.Sleep(1000);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            Console.WriteLine("Document Assemblies have been unloaded.");
            //Console.WriteLine($"Loaded AssemblyCount: {AppDomain.CurrentDomain.GetAssemblies().Length}");
        }

        private static void ExecuteTemplates(out WeakReference loadContextRef)
        {
            var loadCtx = new TestAssemblyLoadContext(Path.GetDirectoryName(typeof(Program).Assembly.Location));
            DocumentFactory.LoadContext = loadCtx;

            var viewPath = $"{BasePath}/Views/Tutorial.cs.docx";
            var documentPath = $"{BasePath}/Documents/Tutorial.docx";
            var imageDirectory = $"{BasePath}/Images";

#if DEBUG
            Ide.Start(viewPath, documentPath, null, null, f => f.ImageDirectory = imageDirectory);

#else
            DocumentFileBase document = DocumentFactory.Create(viewPath);
            document.ImageDirectory = imageDirectory;
            document.Generate(documentPath);
#endif
            loadContextRef = new WeakReference(loadCtx);

            Console.WriteLine("---------------------Assemblies Loaded In the Default Context-------------------------------");
            var assemblyNames = AssemblyLoadContext.Default.Assemblies.Select(s => s.FullName).ToArray();
            Console.WriteLine(string.Join(Environment.NewLine, assemblyNames));

            Console.WriteLine("---------------------Assemblies Loaded In Context-------------------------------");
            assemblyNames = loadCtx.Assemblies.Select(s => s.FullName).ToArray();
            Console.WriteLine(string.Join(Environment.NewLine, assemblyNames));

            loadCtx.Unload();
            DocumentFactory.LoadContext = null;
        }
    }
}