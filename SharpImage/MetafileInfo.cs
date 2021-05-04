using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace SharpImage
{
    public class MetafileInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            var metafile = new Metafile(stream);

            Width = metafile.Width;
            Height = metafile.Height;
            DpiH = (int)metafile.HorizontalResolution;
            DpiV = (int)metafile.VerticalResolution;
            IsValid = true;
        }
    }
}