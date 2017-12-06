using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

        public string ViewPath { get; set; }

        protected abstract void InvokeDocumentCode();

        protected abstract void SetModel(object model);

        public void Generate(string documentPath)
        {
            documentPath = Path.GetFullPath(documentPath);

            File.Copy(ViewPath, documentPath, true);

            using (this.Package = WordprocessingDocument.Open(documentPath, true))
            {
                this.Map = CharacterMap.Create(this.Package.MainDocumentPart.Document.Body);
                this.CodeBlocks = CodeBlockFactory.Create(this.Map, true);

                // Update Map, since we modified the document.
                this.Map = CharacterMap.Create(this.Package.MainDocumentPart.Document.Body);

                InvokeDocumentCode();

                this.Package.Save();
                this.Package.Close();
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
            this.CurrentCodeBlock.Placeholder.Text = ToString(o);
        }

        protected string ToString(object o)
        {
            return o?.ToString() ?? string.Empty;
        }

        protected void DeleteCodeBlock()
        {
            this.Map.Delete(this.CurrentCodeBlock.Placeholder, this.CurrentCodeBlock.EndConditionalPart);
        }

        protected void CreateParagraph()
        {
            this.ParagraphAppender.Append(this.CurrentCodeBlock);
        }

        protected void CreateRow()
        {
            this.RowAppender.Append(this.CurrentCodeBlock);
        }

        protected void Image(string fileName, int percentage = 100)
        {
            var filePath = $"{Path.GetDirectoryName(ViewPath)}\\..\\images\\{fileName}";
            var imageTypePart = ImageHelper.GetImagePartType(filePath);

            const long emusPerTwip = 635;
            var maxWidthInEmus = GetPageContentWidthInTwips() * emusPerTwip;

            Drawing drawing = null;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                drawing = ImageHelper.CreateDrawing(this.Package, fs, imageTypePart, percentage, maxWidthInEmus);
            }

            this.CurrentCodeBlock.Placeholder.InsertAfterSelf(drawing);
        }

        protected long GetPageContentWidthInTwips()
        {
            long width = 10000;

            var sectionProperties = this.Package.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
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