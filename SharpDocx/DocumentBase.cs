using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SharpDocx.Extensions;
using SharpDocx.Models;

namespace SharpDocx
{
    public abstract class DocumentBase
    {
        protected readonly ElementAppender<Paragraph> ParagraphAppender = new ElementAppender<Paragraph>();

        protected readonly ElementAppender<TableRow> RowAppender = new ElementAppender<TableRow>();

        protected List<CodeBlock> CodeBlocks;

        protected CodeBlock CurrentCodeBlock;

        protected CharacterMap Map;

        protected WordprocessingDocument Package;

        public string ImageDirectory { get; set; }

        public string ViewPath { get; private set; }

        protected abstract void InvokeDocumentCode();

        protected abstract void SetModel(object model);

        public void Generate(string documentPath)
        {
            documentPath = Path.GetFullPath(documentPath);

            File.Copy(ViewPath, documentPath, true);

            using (Package = WordprocessingDocument.Open(documentPath, true))
            {
                var codeBlockBuilder = new CodeBlockBuilder(Package, true);
                CodeBlocks = codeBlockBuilder.CodeBlocks;
                Map = codeBlockBuilder.BodyMap;

                InvokeDocumentCode();

                foreach (var cb in CodeBlocks)
                {
                    cb.RemoveEmptyParagraphs();
                }

                Package.Save();
            }
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
            CurrentCodeBlock.Placeholder.Text = ToString(o);
        }

        protected string ToString(object o)
        {
            return o?.ToString() ?? string.Empty;
        }

        protected void Replace(string oldValue, string newValue, int startIndex = 0,
            StringComparison stringComparison = StringComparison.CurrentCulture)
        {
            Map.Replace(oldValue, newValue, startIndex, stringComparison);
        }

        protected void DeleteCodeBlock()
        {
            Map.Delete(CurrentCodeBlock.Placeholder, CurrentCodeBlock.EndConditionalPart);
        }

        protected void AppendParagraph()
        {
            ParagraphAppender.Append(CurrentCodeBlock);
        }

        protected void AppendRow()
        {
            RowAppender.Append(CurrentCodeBlock);
        }

        protected void Image(string filePath, int percentage = 100)
        {
            if (string.IsNullOrEmpty(Path.GetDirectoryName(filePath)) &&
                !string.IsNullOrEmpty(ImageDirectory))
            {
                filePath = $"{ImageDirectory}\\{filePath}";
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
    }
}