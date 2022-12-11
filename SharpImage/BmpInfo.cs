using System.IO;

namespace SharpImage
{
    public class BmpInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            Type = ImageType.Unknown;

            using (var bh = new ByteHelper(stream))
            {
                bh.Seek(0, SeekOrigin.Begin);

                // https://en.wikipedia.org/wiki/BMP_file_format
                IsValid = bh.ReadAscii(2) == "BM";

                if (!IsValid)
                {
                    return;
                }

                bh.Seek(18, SeekOrigin.Begin);
                Width = (int) bh.ReadUint();
                Height = (int) bh.ReadUint();

                bh.Seek(38, SeekOrigin.Begin);
                var pixelsPerMeterH = (int) bh.ReadUint();
                var dpiH = PixelsPerMeterToPixelsPerInch(pixelsPerMeterH);
                if (dpiH > 0)
                {
                    DpiH = dpiH;
                }
                var pixelsPerMeterV = (int) bh.ReadUint();
                var dpiV = PixelsPerMeterToPixelsPerInch(pixelsPerMeterV);
                if (dpiV > 0)
                {
                    DpiV = dpiV;
                }
            }

            Type = ImageType.Bmp;
        }
    }
}