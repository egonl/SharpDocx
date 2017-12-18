//#define DEBUG_DOCUMENT_CODE

using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Security.Permissions;
using System.Text;
using SharpDocx.Models;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.CSharp;

namespace SharpDocx
{
    internal class DocumentCompiler
    {
        public const string Namespace = "SharpDocx";

        private static readonly string documentClassTemplate =
            @"
using System;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Models;
{UsingDirectives}

namespace {Namespace}
{
	public class {ClassName} : {BaseClassName}
	{
        public {ModelType} Model { get; set; }

		protected override void InvokeDocumentCode()
		{
{InvokeDocumentCodeBody}
		}

        protected override void SetModel(object model)
        {
            Model = model as {ModelType};
        }
	}
}";

        internal static Assembly Compile(string viewPath, string className, string baseClassName, object model, List<string> usingDirectives, List<string> referencedAssemblies)
        {
            List<CodeBlock> codeBlocks;

            // Copy the template to a temporary file, so it can be opened even when the template is open in Word.
            // This makes testing templates a lot easier.
            var tempFilePath = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString("d") + ".cs.docx";
            System.IO.File.Copy(viewPath, tempFilePath, true);            

            try
            {
                using (var package = WordprocessingDocument.Open(tempFilePath, false))
                {
                    var codeBlockBuilder = new CodeBlockBuilder(package, false);
                    codeBlocks = codeBlockBuilder.CodeBlocks;
                }
            }
            finally
            {
                System.IO.File.Delete(tempFilePath);
            }

            var invokeDocumentCodeBody = new StringBuilder();

            for (int i = 0; i < codeBlocks.Count; ++i)
            {
                var cb = codeBlocks[i];
                invokeDocumentCodeBody.Append($"            CurrentCodeBlock = CodeBlocks[{i}];{Environment.NewLine}");

                if (!string.IsNullOrEmpty(cb.Code))
                {
                    if (cb.Code[0] == '=')
                    {
                        // Expand <%=SomeVar%> into <% Write(SomeVar); %>
                        invokeDocumentCodeBody.Append($"            Write({cb.Code.Substring(1)});{Environment.NewLine}");
                    }
                    else if (cb.Conditional)
                    {
                        invokeDocumentCodeBody.Append($"if (!{cb.Condition}) {{{Environment.NewLine}");
                        invokeDocumentCodeBody.Append($"DeleteCodeBlock();");
                        invokeDocumentCodeBody.Append($"}}{Environment.NewLine}");

                        invokeDocumentCodeBody.Append(cb.Code);
                    }
                    else
                    {
                        invokeDocumentCodeBody.Append(cb.Code);
                    }
                }

                invokeDocumentCodeBody.Append(Environment.NewLine);
            }

            var modelType = "string";
            if (model != null)
            {
                modelType = FormatType(model.GetType());
            }

            var script = new StringBuilder();
            script.Append(documentClassTemplate);
            script.Replace("{UsingDirectives}", FormatUsingDirectives(usingDirectives));
            script.Replace("{Namespace}", Namespace);
            script.Replace("{ClassName}", className);
            script.Replace("{BaseClassName}", baseClassName);
            script.Replace("{ModelType}", modelType);
            script.Replace("{InvokeDocumentCodeBody}", invokeDocumentCodeBody.ToString());
            return Compile(script.ToString(), referencedAssemblies);
        }

        private static Assembly Compile(
            string sourceCode,
            List<string> referencedAssemblies)
        {
            // Create the compiler.
#if NET35
            var options = new Dictionary<string, string> { { "CompilerVersion", "v3.5" } };
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
                IncludeDebugInformation = false,                
            };

#if DEBUG_DOCUMENT_CODE
            // Create an assembly with debug information and store it in a file. This allows us to step through the generated code.
            // Temporary files are stored in C:\Users\username\AppData\Local\Temp and are not deleted automatically.
            parameters.GenerateInMemory = false;
            parameters.IncludeDebugInformation = true;
            parameters.TempFiles = new TempFileCollection {KeepFiles = true};
            parameters.OutputAssembly = parameters.TempFiles.BasePath + ".dll"; // BasePath is *not* path, but a path with some random filename
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

            // Compile the code.
            var results = compiler.CompileAssemblyFromSource(parameters, sourceCode);
            Debug.WriteLine("***source code***");
            Debug.WriteLine(sourceCode);

            // Raise an SharpDocxCompilationException if errors occured.
            if (results.Errors.HasErrors)
            {
                var formattedCode = new StringBuilder();
                string[] lines = sourceCode.Split('\n');
                for (int i = 0; i < lines.Length; ++i)
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
    }

    [Serializable]
    public class SharpDocxCompilationException : Exception
    {
        public string Errors;
        public string SourceCode;

        public SharpDocxCompilationException(string code, string errors)
        {
            this.SourceCode = code;
            this.Errors = errors;
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