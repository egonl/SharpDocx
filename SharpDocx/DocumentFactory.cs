using System;
using System.Collections;
using System.IO;

namespace SharpDocx
{
    // DocumentFactory purpose:
    // -Generate and cache DocumentAssemblies;
    // -Create instances of a DocumentBase derived type.

    public static class DocumentFactory
    {
        private static readonly Hashtable Assemblies = new Hashtable();
        private static readonly object AssembliesLock = new object();

        public static TBaseClass Create<TBaseClass>(string viewPath, object model = null, bool forceCompile = false)
            where TBaseClass : DocumentBase
        {
            return (TBaseClass) Create(viewPath, model, typeof(TBaseClass), forceCompile);
        }

        public static DocumentBase Create(
            string viewPath,
            object model = null,
            Type baseClassType = null,
            bool forceCompile = false)
        {
            viewPath = Path.GetFullPath(viewPath);

            if (baseClassType == null)
            {
                baseClassType = typeof(DocumentBase);
            }

            var baseClassName = baseClassType.Name;
            var modelTypeName = model?.GetType().Name ?? string.Empty;
            DocumentAssembly da;

            lock (AssembliesLock)
            {
                da = (DocumentAssembly) Assemblies[viewPath + baseClassName + modelTypeName];

                if (da == null || forceCompile)
                {
                    da = new DocumentAssembly(
                        viewPath,
                        baseClassType,
                        model?.GetType());

                    Assemblies[viewPath + baseClassName + modelTypeName] = da;
                }
            }

            var document = (DocumentBase) da.Instance();
            document.Init(viewPath, model);
            return document;
        }
    }
}