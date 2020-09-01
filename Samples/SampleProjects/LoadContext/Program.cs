using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using SharpDocx;

namespace LoadContext
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"/../../../../..";

        static AssemblyLoadContext LoadContext = new TestAssemblyLoadContext(Path.GetDirectoryName(typeof(Program).Assembly.Location));

        private static void Main()
        {
            DocumentFactory.LoadContext = LoadContext;

            var viewPath = $"{BasePath}/Views/Tutorial.cs.docx";
            var documentPath = $"{BasePath}/Documents/Tutorial.docx";
            var imageDirectory = $"{BasePath}/Images";

#if DEBUG
            Ide.Start(viewPath, documentPath, null, null, f => f.ImageDirectory = imageDirectory);

#else
            DocumentBase document = DocumentFactory.Create(viewPath);
            document.ImageDirectory = imageDirectory;
            document.Generate(documentPath);
#endif
            Console.WriteLine("---------------------Assemblies Loaded In the Default Context-------------------------------");
            var assemblyNames = AssemblyLoadContext.Default.Assemblies.Select(s => s.FullName).ToArray();
            Console.WriteLine(string.Join(Environment.NewLine, assemblyNames));

            Console.WriteLine("---------------------Assemblies Loaded In Context-------------------------------");
             assemblyNames = LoadContext.Assemblies.Select(s => s.FullName).ToArray();
            Console.WriteLine(string.Join(Environment.NewLine, assemblyNames));

            LoadContext.Unload();
            
            Console.WriteLine("Document Assemblies have been unloaded.");
        }

        class TestAssemblyLoadContext : AssemblyLoadContext
        {
            private readonly AssemblyDependencyResolver _resolver;

            public TestAssemblyLoadContext(string mainAssemblyToLoadPath) : base(isCollectible: true)
            {
                _resolver = new AssemblyDependencyResolver(mainAssemblyToLoadPath);
            }

            protected override Assembly Load(AssemblyName name)
            {
                string assemblyPath = _resolver.ResolveAssemblyToPath(name);
                if (assemblyPath != null)
                {
                    return LoadFromAssemblyPath(assemblyPath);
                }

                return null;
            }
        }
    }
}