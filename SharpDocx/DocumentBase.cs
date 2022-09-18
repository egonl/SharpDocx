//#define SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using SharpDocx.CodeBlocks;
using SharpDocx.Extensions;
using System.Diagnostics;

namespace SharpDocx
{
    public abstract class DocumentBase
    {
        public string ImageDirectory { get; set; }

        public string ViewPath { get; private set; }

        protected List<CodeBlock> CodeBlocks;

        protected CodeBlock CurrentCodeBlock;

        protected TextBlock CurrentTextBlock = null;

        protected Stack<TextBlock> CurrentTextBlockStack = new Stack<TextBlock>();

        protected CharacterMap Map;

        protected WordprocessingDocument Package;

        protected abstract void InvokeDocumentCode();

        public abstract void SetModel(object model);

#if NET35 && SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35
        private static readonly Mutex PackageMutex = new Mutex(false);
#endif

        public void Generate(string documentPath, object model = null)
        {
            documentPath = Path.GetFullPath(documentPath);
            File.Copy(ViewPath, documentPath, true);

#if NET35 && SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35
            // Due to a bug in System.IO.Packaging writing large uncompressed parts (>10MB) isn't thread safe in .NET 3.5.
            // Workaround: make writing in all threads and processes sequential.
            // Microsoft fixed this in .NET 4.5 (see https://maheshkumar.wordpress.com/2014/10/21/).
            PackageMutex.WaitOne(Timeout.Infinite, false);

            try
            {
#endif
                using (Package = WordprocessingDocument.Open(documentPath, true))
                {
                    GenerateInternal(model);
                }

#if NET35 && SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35
            }
            finally
            {
                PackageMutex.ReleaseMutex();
            }
#endif
        }

        public MemoryStream Generate(object model = null)
        {
            var viewBytes = File.ReadAllBytes(ViewPath);
            MemoryStream outputstream = new MemoryStream();
            outputstream.Write(viewBytes, 0, viewBytes.Length);

#if NET35 && SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35
            // Due to a bug in System.IO.Packaging writing large uncompressed parts (>10MB) isn't thread safe in .NET 3.5.
            // Workaround: make writing in all threads and processes sequential.
            // Microsoft fixed this in .NET 4.5 (see https://maheshkumar.wordpress.com/2014/10/21/).
            PackageMutex.WaitOne(Timeout.Infinite, false);

            try
            {
#endif
                using (Package = WordprocessingDocument.Open(outputstream, true))
                {
                    GenerateInternal(model);
                }

#if NET35 && SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35
            }
            finally
            {
                PackageMutex.ReleaseMutex();
            }
#endif
            // Reset position of stream.
            outputstream.Seek(0, SeekOrigin.Begin);
            return outputstream;
        }

        protected void GenerateInternal(object model)
        {
            if (model != null)
            {
                SetModel(model);
            }

            var codeBlockBuilder = new CodeBlockBuilder(Package);
            CodeBlocks = codeBlockBuilder.CodeBlocks;
            Map = codeBlockBuilder.BodyMap;

            InvokeDocumentCode();

            foreach (var cb in CodeBlocks)
            {
                cb.RemoveEmptyParagraphs();
            }

            CreateUniqueDocPropertiesIds();
        }

        // Override this static method if you want to specify additional using directives.
        public static List<string> GetUsingDirectives()
        {
            return null;
        }

        // Override this static method if you want to reference certain assemblies.
        public static List<string> GetReferencedAssemblies()
        {
            return null;
        }

        internal void Init(string viewPath, object model)
        {
            ViewPath = viewPath;
            SetModel(model);
        }

        protected void Write(object o)
        {
            var s = ToString(o);

            var lines = s.Split('\n');
            if (lines.Length == 1)
            {
                CurrentCodeBlock.Placeholder.Text = s;
                return;
            }

            CurrentCodeBlock.Placeholder.Text = lines[0];
            var lastText = CurrentCodeBlock.Placeholder;

            for (var i = 1; i < lines.Length; ++i)
            {
                var br = lastText.InsertAfterSelf(new Break());
                lastText = br.InsertAfterSelf(new Text(lines[i]));
            }
        }

        protected string ToString(object o)
        {
            return o?.ToString() ?? string.Empty;
        }

#if NET35
        protected void Replace(string oldValue, string newValue)
        {
            Replace(oldValue, newValue, 0, StringComparison.CurrentCulture);
        }
#endif

        protected void Replace(string oldValue, string newValue, int startIndex = 0,
            StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            Map.Replace(oldValue, newValue, startIndex, stringComparison);
        }

        protected void AppendTextBlock()
        {
            CurrentTextBlock.Append(CodeBlocks);
        }

        protected void AppendParagraph()
        {
            var pa = CurrentCodeBlock as ParagraphAppender;
            pa.Append();
        }

        protected void AppendRow()
        {
            var ra = CurrentCodeBlock as RowAppender;
            ra.Append();
        }

#if NET35
        protected void Image(string filePath)
        {
            Image(filePath, 100);
        }
#endif

        protected void Image(string filePath, int percentage = 100)
        {
            if (string.IsNullOrEmpty(Path.GetDirectoryName(filePath)) &&
                !string.IsNullOrEmpty(ImageDirectory))
            {
                filePath = $"{ImageDirectory}/{filePath}";
            }

            if (!File.Exists(filePath))
            {
#if DEBUG
                CurrentCodeBlock.Placeholder.Text = $"Image '{filePath}' not found.";
#endif
                return;
            }

            var imageTypePart = ImageHelper.GetImagePartType(filePath);

            const long emusPerTwip = 635;
            var maxWidthInEmus = GetPageContentWidthInTwips() * emusPerTwip;

            Drawing drawing;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                drawing = ImageHelper.CreateDrawing(Package, fs, imageTypePart, percentage, maxWidthInEmus);
            }

            CurrentCodeBlock.Placeholder.InsertAfterSelf(drawing);

            if (!CurrentCodeBlock.Placeholder.GetParent<Paragraph>().HasText())
            {
                // Insert a zero-width space, so the image doesn't get deleted by CodeBlock.RemoveEmptyParagraphs.
                CurrentCodeBlock.Placeholder.Text = "\u200B";
            }
        }

        protected long GetPageContentWidthInTwips()
        {
            long width = 10000;

            var sectionProperties = Package.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
            var pageSize = sectionProperties?.GetFirstChild<PageSize>();
            if (pageSize != null)
            {
                var pageMargin = sectionProperties.GetFirstChild<PageMargin>();

                if (pageMargin != null)
                {
                    width = pageSize.Width - pageMargin.Left - pageMargin.Right;
                }
                else
                {
                    width = pageSize.Width - 1800;
                }
            }

            return width;
        }

        private void CreateUniqueDocPropertiesIds()
        {
            // Duplicate DocProperties.Id might lead to issues when opening the file in Word, see https://github.com/python-openxml/python-docx/issues/455#issuecomment-357748310.
            // Here we'll de-duplicate any duplicates.

            var docProperties = GetDocProperties();
            var usedIds = new List<uint>();

            foreach (var docProperty in docProperties)
            {
                //Debug.Write($"ID was {docProperty.Id}");

                while (usedIds.Contains(docProperty.Id))
                {
                    ++docProperty.Id;
                }

                usedIds.Add(docProperty.Id);

                //Debug.WriteLine($" and is now {docProperty.Id}");
            }
        }

        private List<DocProperties> GetDocProperties()
        {
            var list = new List<DocProperties>();
            List<DocProperties> partialResult;

            foreach (var headerPart in Package.MainDocumentPart.HeaderParts)
            {
                partialResult = headerPart.Header.GetAllElements<DocProperties>();
                list.AddRange(partialResult);
            }

            foreach (var footerPart in Package.MainDocumentPart.FooterParts)
            {
                partialResult = footerPart.Footer.GetAllElements<DocProperties>();
                list.AddRange(partialResult);
            }

            if (Package.MainDocumentPart.EndnotesPart != null)
            {
                partialResult = Package.MainDocumentPart.EndnotesPart.Endnotes.GetAllElements<DocProperties>();
                list.AddRange(partialResult);
            }

            if (Package.MainDocumentPart.FootnotesPart != null)
            {
                partialResult = Package.MainDocumentPart.FootnotesPart.Footnotes.GetAllElements<DocProperties>();
                list.AddRange(partialResult);
            }

            partialResult = Package.MainDocumentPart.Document.Body.GetAllElements<DocProperties>();
            list.AddRange(partialResult);

            return list;
        }


    }
}