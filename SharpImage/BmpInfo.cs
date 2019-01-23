using System.IO;

namespace SharpImage
{
    public class BmpInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
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
                DpiH = PixelsPerMeterToPixelsPerInch(pixelsPerMeterH);
                var pixelsPerMeterV = (int) bh.ReadUint();
                DpiV = PixelsPerMeterToPixelsPerInch(pixelsPerMeterV);
            }
        }
    }
}