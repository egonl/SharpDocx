using System.IO;

namespace SharpImage
{
    public class PngInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            using (var bh = new ByteHelper(stream))
            {
                bh.IsLsbf = false;
                bh.Seek(0, SeekOrigin.Begin);

                // http://www.libpng.org/pub/png/spec/1.2/PNG-Structure.html
                var signature = bh.ReadBytes(8);

                IsValid = signature[0] == 137 && signature[1] == 80 && signature[2] == 78 && signature[3] == 71 &&
                          signature[4] == 13 && signature[5] == 10 && signature[6] == 26 && signature[7] == 10;

                if (!IsValid)
                {
                    return;
                }

                DpiH = DpiV = 96;

                while (stream.Position < stream.Length)
                {
                    var chunkDataLength = bh.ReadUint();
                    var chunkType = bh.ReadAscii(4);

                    switch (chunkType)
                    {
                        case "IHDR":
                            Width = (int) bh.ReadUint();
                            Height = (int) bh.ReadUint();
                            bh.Seek(5);
                            break;

                        case "pHYs":
                            var pixelsPerUnitX = (int) bh.ReadUint();
                            var pixelsPerUnitY = (int) bh.ReadUint();
                            var unit = bh.ReadByte();
                            if (unit == 1)
                            {
                                // Unit is the meter.
                                DpiH = PixelsPerMeterToPixelsPerInch(pixelsPerUnitX);
                                DpiV = PixelsPerMeterToPixelsPerInch(pixelsPerUnitY);
                            }

                            return;

                        default:
                            bh.Seek(chunkDataLength, SeekOrigin.Current);
                            break;
                    }

                    // Skip CRC.
                    bh.Seek(4, SeekOrigin.Current);
                }
            }
        }
    }
}