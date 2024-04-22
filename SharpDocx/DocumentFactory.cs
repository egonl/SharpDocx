using System;
using System.Collections;
using System.IO;
using System.Security.Cryptography;
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

            if (!File.Exists(viewPath))
            {
                throw new ArgumentException($"Could not find the file '{viewPath}'", nameof(viewPath));
            }

            var viewLastWriteTime = new FileInfo(viewPath).LastWriteTime.ToString("s", System.Globalization.CultureInfo.InvariantCulture);
            var viewId = viewPath + viewLastWriteTime;

#if DEBUG
            // Copy the template to a temporary file, so it can be opened even when the template is open in Word.
            // This makes testing templates a lot easier.
            var tempFilePath = $"{Path.GetTempPath()}{Guid.NewGuid():N}.cs.docx";
            File.Copy(viewPath, tempFilePath, true);
            using var viewStream = File.OpenRead(tempFilePath);

            DocumentBase documentBase;
            try
            {
                documentBase = CreateInternal(viewId, viewStream, model, baseClassType, forceCompile);
            }
            finally
            {
                viewStream.Close();
                File.Delete(tempFilePath);
            }
#else
            using var viewStream = File.OpenRead(viewPath);
            var documentBase = CreateInternal(viewId, viewStream, model, baseClassType, forceCompile);
#endif

            documentBase.ViewPath = viewPath;
            return documentBase;
        }

        public static TBaseClass Create<TBaseClass>(Stream viewStream, object model = null, bool forceCompile = false)
            where TBaseClass : DocumentBase
        {
            return (TBaseClass)Create(viewStream, model, typeof(TBaseClass), forceCompile);
        }

        public static DocumentBase Create(
            Stream viewStream,
            object model = null,
            Type baseClassType = null,
            bool forceCompile = false)
        {
            var viewId = GetHash(viewStream);
            var documentBase = CreateInternal(viewId, viewStream, model, baseClassType, forceCompile);
            documentBase.ViewStream = viewStream;
            return documentBase;
        }

        private static DocumentBase CreateInternal(
            string viewId,
            Stream viewStream,
            object model = null,
            Type baseClassType = null,
            bool forceCompile = false)
        {
            if (baseClassType == null)
            {
                baseClassType = typeof(DocumentBase);
            }

            var baseClassName = baseClassType.Name;
            var modelTypeName = model?.GetType().Name ?? string.Empty;
            DocumentAssembly da;

            lock (AssembliesLock)
            {
                da = (DocumentAssembly)Assemblies[viewId + baseClassName + modelTypeName];

                if (da == null || forceCompile)
                {
                    da = new DocumentAssembly(
                        viewStream,
                        baseClassType,
                        model?.GetType());

                    Assemblies[viewId + baseClassName + modelTypeName] = da;
                }
            }

            var document = (DocumentBase)da.Instance();
            document.Init(model);
            return document;
        }

        private static string GetHash(Stream s)
        {
            using SHA256 sha256 = SHA256.Create();
            s.Seek(0, SeekOrigin.Begin);
            var hashBytes = sha256.ComputeHash(s);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
        }
    }
}