//#define SUPPORT_MULTI_THREADING_AND_LARGE_DOCUMENTS_IN_NET35

using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace SharpDocx
{
    public abstract class DocumentFileBase : DocumentBase
    {
        public string ViewPath { get; private set; }

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

        internal void Init(string viewPath, object model)
        {
            ViewPath = viewPath;
            base.InitBase(model);
        }
    }
}