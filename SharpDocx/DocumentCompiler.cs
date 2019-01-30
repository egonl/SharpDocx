//#define DEBUG_DOCUMENT_CODE

#if !(NET35 || NET45)
#define AUTO_REFERENCE_SDK
#endif

#if !NET35
using System.Linq;
#endif

#if NET35 || NET45
using System.CodeDom.Compiler;
using Microsoft.CSharp;
#else
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using SharpImage;
#endif

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using SharpDocx.CodeBlocks;

namespace SharpDocx
{
    internal class DocumentCompiler
    {
        public const string Namespace = "SharpDocx";

        private static readonly string documentClassTemplate =
            @"using System;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.CodeBlocks;
using SharpDocx.Models;
{UsingDirectives}

namespace {Namespace}
{
    public class {ClassName} : {BaseClassName}
    {
        public {ModelTypeName} Model { get; set; }

        protected override void InvokeDocumentCode()
        {
{InvokeDocumentCodeBody}
        }

        public override void SetModel(object model)
        {
            if (model == null)
            {
                Model = null;
                return;
            }

            Model = model as {ModelTypeName};

            if (Model == null)
            {
                throw new ArgumentException(""Model is not of type {ModelTypeName}"", ""model"");
            }
        }
    }
}";

        internal static Assembly Compile(
            string viewPath, 
            string className, 
            string baseClassName, 
            Type modelType,
            List<string> usingDirectives, 
            List<string> referencedAssemblies)
        {
            List<CodeBlock> codeBlocks;

            // Copy the template to a temporary file, so it can be opened even when the template is open in Word.
            // This makes testing templates a lot easier.
            var tempFilePath = $"{Path.GetTempPath()}{Guid.NewGuid():N}.cs.docx";
            File.Copy(viewPath, tempFilePath, true);

            try
            {
                using (var package = WordprocessingDocument.Open(tempFilePath, false))
                {
                    var codeBlockBuilder = new CodeBlockBuilder(package);
                    codeBlocks = codeBlockBuilder.CodeBlocks;
                }
            }
            finally
            {
                File.Delete(tempFilePath);
            }

            var invokeDocumentCodeBody = new StringBuilder();
            Stack<TextBlock> currentTextBlockStack = new Stack<TextBlock>();

            for (var i = 0; i < codeBlocks.Count; ++i)
            {
                var cb = codeBlocks[i];

                var tb = cb as TextBlock;
                if (tb != null)
                {
                    currentTextBlockStack.Push(tb);
                    invokeDocumentCodeBody.Append($"            CurrentTextBlockStack.Push(CurrentTextBlock);{Environment.NewLine}");
                    invokeDocumentCodeBody.Append($"            CurrentTextBlock = CodeBlocks[{i}] as TextBlock;{Environment.NewLine}");
                }

                if (!string.IsNullOrEmpty(cb.Code))
                {
                    if (cb.Code[0] == '=')
                    {
                        // Expand <%=SomeVar%> into <% Write(SomeVar); %>
                        invokeDocumentCodeBody.Append($"            CurrentCodeBlock = CodeBlocks[{i}];{Environment.NewLine}");
                        invokeDocumentCodeBody.Append($"            Write({cb.Code.Substring(1)});{Environment.NewLine}");
                    }
                    else if (cb is Directive)
                    {
                        var directive = (Directive) cb;
                        if (directive.Name.Equals("import"))
                        {
                            AddUsingDirective(directive, usingDirectives);
                        }
                        else if (directive.Name.Equals("assembly"))
                        {
                            AddReferencedAssembly(directive, referencedAssemblies);
                        }
                    }
                    else
                    {
                        if (currentTextBlockStack.Count > 0 && cb == currentTextBlockStack.Peek().EndingCodeBlock)
                    {
                            // Automatically insert AppendTextBlock before closing text block brace.
                            invokeDocumentCodeBody.Append($"            AppendTextBlock();{Environment.NewLine}");
                    }
                        invokeDocumentCodeBody.Append($"            CurrentCodeBlock = CodeBlocks[{i}];{Environment.NewLine}");
                        invokeDocumentCodeBody.Append($"            {cb.Code.TrimStart()}{Environment.NewLine}");
                    }
                }

                if (currentTextBlockStack.Count > 0 && cb == currentTextBlockStack.Peek().EndingCodeBlock)
                {
                    currentTextBlockStack.Pop();
                    invokeDocumentCodeBody.Append($"            CurrentTextBlock = CurrentTextBlockStack.Pop();{Environment.NewLine}");
                }
            }

            var modelTypeName = modelType != null ? FormatType(modelType) : "string";

            var script = new StringBuilder();
            script.Append(documentClassTemplate);
            script.Replace("{UsingDirectives}", FormatUsingDirectives(usingDirectives));
            script.Replace("{Namespace}", Namespace);
            script.Replace("{ClassName}", className);
            script.Replace("{BaseClassName}", baseClassName);
            script.Replace("{ModelTypeName}", modelTypeName);
            script.Replace("{InvokeDocumentCodeBody}", invokeDocumentCodeBody.ToString());
            return Compile(script.ToString(), referencedAssemblies);
        }

        private static void AddUsingDirective(Directive directive, List<string> usingDirectives)
        {
            // Support for <%@ Import Namespace="System.Collections.Generic" %>
            if (!directive.Attributes.ContainsKey("namespace"))
            {
                throw new Exception($"The Import directive requires a Namespace attribute in '{directive.Code}'.");
            }

            usingDirectives.Add($"using {directive.Attributes["namespace"]};");
        }

        private static void AddReferencedAssembly(Directive directive, List<string> referencedAssemblies)
        {
            // Support for <%@ Assembly Name="System.Speech" %>
            if (!directive.Attributes.ContainsKey("name"))
            {
                throw new Exception($"The Assembly directive requires a Name attribute in '{directive.Code}'.");
            }

            var assembly = directive.Attributes["name"];
            if (!assembly.EndsWith(".dll", StringComparison.InvariantCultureIgnoreCase) &&
                !assembly.EndsWith(".exe", StringComparison.InvariantCultureIgnoreCase))
            {
                assembly = assembly + ".dll";
            }

            referencedAssemblies.Add(assembly);
        }

        private static Assembly Compile(
            string sourceCode,
            List<string> referencedAssemblies)
        {
            Debug.WriteLine("***Source code***");
            Debug.WriteLine(sourceCode);

#if NET35 || NET45
            // Create the compiler.
#if NET35
            var options = new Dictionary<string, string> {{"CompilerVersion", "v3.5"}};
#else
            var options = new Dictionary<string, string> { { "CompilerVersion", "v4.0" } };
#endif
            CodeDomProvider compiler = new CSharpCodeProvider(options);

            // Add compiler options.
            var parameters = new CompilerParameters
            {
                CompilerOptions = "/target:library",
                GenerateExecutable = false,
                GenerateInMemory = true,
                IncludeDebugInformation = false
            };

#if DEBUG_DOCUMENT_CODE 
            // Create an assembly with debug information and store it in a file. This allows us to step through the generated code.
            // Temporary files are stored in C:\Users\username\AppData\Local\Temp and are not deleted automatically.
            parameters.GenerateInMemory = false;
            parameters.IncludeDebugInformation = true;
            parameters.TempFiles = new TempFileCollection {KeepFiles = true};
            parameters.OutputAssembly = $"{parameters.TempFiles.BasePath}.dll";
#else 
            // Create an assembly in memory and do not include debug information. Fast, but you can't step through the code.
            parameters.CompilerOptions = "/target:library /optimize";
#endif
            // Add referenced assemblies.
            parameters.ReferencedAssemblies.Add("mscorlib.dll");
            parameters.ReferencedAssemblies.Add("System.dll");
            parameters.ReferencedAssemblies.Add(typeof(WordprocessingDocument).Assembly.Location);
            parameters.ReferencedAssemblies.Add(Assembly.GetExecutingAssembly().Location);

            if (referencedAssemblies != null)
            {
                foreach (var ra in referencedAssemblies)
                {
                    parameters.ReferencedAssemblies.Add(ra);
                }
            }

#if DEBUG
            Debug.WriteLine("***References assemblies");
            foreach (var reference in parameters.ReferencedAssemblies)
            {
                Debug.WriteLine(reference);
            }
#endif

            // Compile the code.
            var results = compiler.CompileAssemblyFromSource(parameters, sourceCode);

            // Raise an SharpDocxCompilationException if errors occured.
            if (results.Errors.HasErrors)
            {
                var formattedCode = new StringBuilder();
                var lines = sourceCode.Split('\n');
                for (var i = 0; i < lines.Length; ++i)
                {
                    formattedCode.Append($"{i + 1,5}  {lines[i]}\n");
                }

                var formattedErrors = new StringBuilder();
                foreach (CompilerError e in results.Errors)
                {
                    // Do not show the name of the temporary file.
                    e.FileName = "";
                    formattedErrors.Append($"Line {e.Line}: {e}{Environment.NewLine}{Environment.NewLine}");
                }

                throw new SharpDocxCompilationException(formattedCode.ToString(), formattedErrors.ToString());
            }

            // Return the assembly.
            return results.CompiledAssembly;
#else
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(sourceCode);

            // Get the path to the shared assemblies, e.g. C:\Program Files\dotnet\shared\Microsoft.NETCore.App\2.0.9.
            var assemblyDir = Path.GetDirectoryName(typeof(object).Assembly.Location);

            var references = new List<MetadataReference>
            {
                MetadataReference.CreateFromFile(Path.Combine(assemblyDir, "mscorlib.dll")),
                MetadataReference.CreateFromFile(Path.Combine(assemblyDir, "netstandard.dll")),
                MetadataReference.CreateFromFile(Assembly.GetExecutingAssembly().Location),
                MetadataReference.CreateFromFile(typeof(WordprocessingDocument).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(DocumentBase).Assembly.Location)
            };

#if AUTO_REFERENCE_SDK
            // Auto reference all managed Microsoft- or System-DLL's.
            foreach (var dllPath in Directory.GetFiles(assemblyDir, "*.dll"))
            {
                var fileName = Path.GetFileName(dllPath);
                if (fileName.StartsWith("Microsoft.") || fileName.StartsWith("System."))
                {
                    using (var stream = File.OpenRead(dllPath))
                    {
                        if (IsManagedDll(stream))
                        {
                            references.Add(MetadataReference.CreateFromFile(dllPath));
                        }
                    }
                }
            }
#else
            // Only reference the bare minimum required DLL's. This saves around 10MB when running.
            references.Add(MetadataReference.CreateFromFile(typeof(object).Assembly.Location));
            references.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyDir, "System.Runtime.dll")));
            references.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyDir, "System.Collections.dll")));
#endif

            if (referencedAssemblies != null)
            {
                foreach (var ra in referencedAssemblies)
                {
                    if (ra.Contains("\\") || ra.Contains("/"))
                    {
                        references.Add(MetadataReference.CreateFromFile(ra));
                    }
                    else
                    {
                        references.Add(MetadataReference.CreateFromFile(Path.Combine(assemblyDir,ra)));
                    }
                }
            }

#if DEBUG
            Debug.WriteLine("***References assemblies");
            foreach (var reference in references.OrderBy(r => r.Display))
            {
                Debug.WriteLine(reference.Display);
            }
#endif

            CSharpCompilation compilation = CSharpCompilation.Create(
                $"SharpAssembly_{Guid.NewGuid():N}",
                new[] { syntaxTree },
                references,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            using (var ms = new MemoryStream())
            {
                EmitResult result = compilation.Emit(ms);

                if (!result.Success)
                {
                    var formattedCode = new StringBuilder();
                    var lines = sourceCode.Split('\n');
                    for (var i = 0; i < lines.Length; ++i)
                    {
                        formattedCode.Append($"{i + 1,5}  {lines[i]}\n");
                    }

                    IEnumerable<Diagnostic> failures = result.Diagnostics.Where(diagnostic => 
                        diagnostic.IsWarningAsError || 
                        diagnostic.Severity == DiagnosticSeverity.Error);

                    var formattedErrors = new StringBuilder();
                    foreach (Diagnostic diagnostic in failures)
                    {
                        // TODO: show line number.
                        formattedErrors.AppendLine($"{diagnostic.Id}: {diagnostic.GetMessage()}");
                    }

                    throw new SharpDocxCompilationException(formattedCode.ToString(), formattedErrors.ToString());
                }

                ms.Seek(0, SeekOrigin.Begin);
                return Assembly.Load(ms.ToArray());
            }
#endif
        }

        private static string FormatUsingDirectives(List<string> usingDirectives)
        {
            if (usingDirectives == null)
            {
                return null;
            }

            var sb = new StringBuilder();

            foreach (var @namespace in usingDirectives)
            {
                sb.Append(@namespace);
                sb.Append(Environment.NewLine);
            }

            return sb.ToString();
        }

        private static string FormatType(Type type)
        {
#if !NET35
            if (type.IsConstructedGenericType)
            {
                var name = type.Name.Substring(0, type.Name.IndexOf("`"));
                var args = type.GenericTypeArguments.Select(FormatType);
                return $"{name}<{string.Join(", ", args)}>";
            }
#endif
            return type.Name;
        }

#if AUTO_REFERENCE_SDK
        public static bool IsManagedDll(Stream dll)
        {
            // See https://docs.microsoft.com/en-us/windows/desktop/debug/pe-format#file-headers

            using (var bh = new ByteHelper(dll))
            {
                bh.Seek(0, SeekOrigin.Begin);

                // Check DOS header.
                var dosSignature = bh.ReadUshort();
                if (dosSignature != 0x5A4D)
                {
                    return false;
                }

                // Read location of PE signature.
                bh.Seek(0x3C, SeekOrigin.Begin);
                var peHeaderOffset = bh.ReadUint();
                
                // Check PE signature.
                bh.Seek(peHeaderOffset, SeekOrigin.Begin);
                var peSignature = bh.ReadUint();
                if (peSignature != 0x4550)
                {
                    return false;
                }

                // Seek to and check optional header.
                bh.Offset = peHeaderOffset + 24;
                bh.Seek(0, SeekOrigin.Begin);
                var optionalHeaderMagic = bh.ReadUshort();
                bool isPe32Plus;
                if (optionalHeaderMagic == 0x10B)
                {
                    isPe32Plus = false;
                }
                else if (optionalHeaderMagic == 0x20B)
                {
                    isPe32Plus = true;
                }
                else
                {
                    return false;
                }

                // Seek to and check NumberOfRvaAndSizes.
                bh.Seek(isPe32Plus ? 108 : 92, SeekOrigin.Begin);
                var numberOfRvaAndSizes = bh.ReadUint();
                if (numberOfRvaAndSizes < 15)
                {
                    return false;
                }

                // Seek to and check CLR Runtime Header data directory.
                bh.Seek(isPe32Plus ? 224 : 208, SeekOrigin.Begin);
                var address = bh.ReadUint();
                var size = bh.ReadUint();
                return address != 0 && size != 0;
            }
        }
#endif
    }

    [Serializable]
    public class SharpDocxCompilationException : Exception
    {
        public string Errors;
        public string SourceCode;

        public SharpDocxCompilationException(string code, string errors)
        {
            SourceCode = code;
            Errors = errors;
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(nameof(Errors), Errors);
            info.AddValue(nameof(SourceCode), SourceCode);
        }
    }
}