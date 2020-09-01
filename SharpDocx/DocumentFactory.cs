using System;
using System.Collections;
using System.IO;
#if NETSTANDARD2_0
using System.Runtime.Loader;
#endif

namespace SharpDocx
{
    // DocumentFactory purpose:
    // -Generate and cache DocumentAssemblies;
    // -Create instances of a DocumentBase derived type.

    public static class DocumentFactory
    {
        private static readonly Hashtable Assemblies = new Hashtable();
        private static readonly object AssembliesLock = new object();

#if NETSTANDARD2_0

        private static readonly object LoadContextLock = new object();
        private static AssemblyLoadContext _loadContext;

        /// <summary>
        /// Sets the load context for the generated assemblies. Setting this context or calling <see cref="AssemblyLoadContext"/> will clear assembly cache.
        /// <para></para>
        /// </summary>
        public static AssemblyLoadContext LoadContext
        {
            get
            {
                lock (LoadContextLock)
                {
                    return _loadContext;
                }
            }
            set
            {
                lock (LoadContextLock)
                {
                    //if replacing load context clear assemblies
                    if (_loadContext != null && _loadContext != value)
                    {
                        lock (AssembliesLock)
                        {
                            Assemblies.Clear();
                        }

                        _loadContext.Unloading -= LoadContextOnUnloading;
                    }

                    _loadContext = value;

                    if (_loadContext != null)
                    {
                        //any call to unload should clear assemblies as the load context will be unusable
                        _loadContext.Unloading += LoadContextOnUnloading;
                    }
                }
            }
        }

        private static void LoadContextOnUnloading(AssemblyLoadContext obj)
        {
            lock (LoadContextLock)
            {
                lock (AssembliesLock)
                {
                    Assemblies.Clear();
                }
            }
        }

#endif

        public static TBaseClass Create<TBaseClass>(string viewPath, object model = null, bool forceCompile = false)
            where TBaseClass : DocumentBase
        {
            return (TBaseClass)Create(viewPath, model, typeof(TBaseClass), forceCompile);
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
                da = (DocumentAssembly)Assemblies[viewPath + baseClassName + modelTypeName];

                if (da == null || forceCompile)
                {
                    da = new DocumentAssembly(
                        viewPath,
                        baseClassType,
                        model?.GetType());

                    Assemblies[viewPath + baseClassName + modelTypeName] = da;
                }
            }

            var document = (DocumentBase)da.Instance();
            document.Init(viewPath, model);
            return document;
        }
    }
}