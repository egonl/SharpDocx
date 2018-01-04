using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using SharpDocx.Extensions;

namespace SharpDocx
{
    internal class DocumentAssembly
    {
        private readonly Assembly _assembly;
        private readonly string _className;

        internal DocumentAssembly(
            string viewPath,
            Type baseClass,
            Type modelType)
        {
            if (!File.Exists(viewPath))
            {
                throw new ArgumentException($"Could not find the file '{viewPath}'", nameof(viewPath));
            }

            if (baseClass == null)
            {
                throw new ArgumentNullException(nameof(baseClass));                
            }

            // Load base class assembly.
            var a = Assembly.LoadFrom(baseClass.Assembly.Location);
            if (a == null)
            {
                throw new ArgumentException($"Can't load assembly '{baseClass.Assembly}'", nameof(baseClass));
            }

            // Get the base class type.
            var t = a.GetType(baseClass.FullName);
            if (t == null)
            {
                throw new ArgumentException(
                    $"Can't find base class '{baseClass.FullName}' in assembly '{baseClass.Assembly}'",
                    nameof(baseClass));
            }

            // Check base class type.
            if (t != typeof(DocumentBase) && !t.IsSubclassOf(typeof(DocumentBase)))
            {
                throw new ArgumentException("baseClass should be a DocumentBase derived type", nameof(baseClass));
            }

            // Get user defined using directives by calling the static DocumentBase.GetUsingDirectives method.
            var usingDirectives =
                (List<string>) a.Invoke(
                    baseClass.FullName,
                    null,
                    "GetUsingDirectives",
                    null)
                ?? new List<string>();
            

            // Get user defined assemblies to reference.
            var referencedAssemblies =
                (List<string>) a.Invoke(
                    baseClass.FullName,
                    null,
                    "GetReferencedAssemblies",
                    null)
                ?? new List<string>();

            if (modelType != null)
            {
                // Add namespace(s) of Model and reference Model assembly/assemblies.
                foreach (var type in GetTypes(modelType))
                {
                    usingDirectives.Add($"using {type.Namespace};");
                    referencedAssemblies.Add(type.Assembly.Location);
                }
            }

            // Create a unique class name.
            _className = $"SharpDocument_{Guid.NewGuid():N}";

            // Create an assembly for this class.
            _assembly = DocumentCompiler.Compile(
                viewPath,
                _className,
                baseClass.Name,
                modelType,
                usingDirectives,
                referencedAssemblies);
        }

        public object Instance()
        {
            return _assembly.CreateInstance($"{DocumentCompiler.Namespace}.{_className}", null);
        }

        private static IEnumerable<Type> GetTypes(Type type)
        {
#if !NET35
            if (type.IsConstructedGenericType)
            {
                foreach (var t in type.GenericTypeArguments)
                {
                    yield return t;
                }
            }
#endif
            yield return type;
        }
    }
}