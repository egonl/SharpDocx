using System;
using System.IO;

namespace SharpImage
{
    public class JpegInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            using (var bh = new ByteHelper(stream))
            {
                bh.Seek(0, SeekOrigin.Begin);
                bh.IsLsbf = false;

                // Start of Image (SOI) marker.
                var marker = bh.ReadBytes(2);
                if (marker[0] != 0xFF && marker[1] != 0xD8)
                {
                    return;
                }

                DpiH = DpiV = 96;
                int length;

                while (stream.Position < stream.Length)
                {
                    marker = bh.ReadBytes(2);
                    if (marker[0] != 0xFF)
                    {
                        throw new Exception("Invalid JPEG file.");
                    }

                    switch (marker[1])
                    {
                        case 0xE0:
                            // APP0 (JFIF) marker.
                            length = bh.ReadUshort();
                            bh.Seek(7);
                            var densityUnits = bh.ReadByte();
                            var xdensity = bh.ReadUshort();
                            var ydensity = bh.ReadUshort();
                            if (densityUnits == 1)
                            {
                                DpiH = xdensity;
                                DpiV = ydensity;
                            }
                            else if (densityUnits == 2)
                            {
                                // Pixels per centimeter
                                DpiH = PixelsPerMeterToPixelsPerInch(xdensity * 100);
                                DpiV = PixelsPerMeterToPixelsPerInch(ydensity * 100);
                            }

                            bh.Seek(length - 14);
                            break;

                        //case 0xE1: 
                        // APP1-marker. This could be an Exif marker.
                        // TODO: read x/y-resolution from Exif APP1

                        case 0xFF:
                            // Filler byte.
                            bh.Seek(-1, SeekOrigin.Current);
                            break;

                        case 0xD0:
                        case 0xD1:
                        case 0xD2:
                        case 0xD3:
                        case 0xD4:
                        case 0xD5:
                        case 0xD6:
                        case 0xD7:
                        case 0xD8:
                        case 0xD9:
                            //  RSTn are used for resync, may be ignored.
                            break;

                        case 0xDA:
                            // SOS (Start of Scan) marker, followed by image data.
                            IsValid = true;
                            return;

                        case 0xC0:
                        case 0xC1:
                        case 0xC2:
                        case 0xC3:
                        case 0xC5:
                        case 0xC6:
                        case 0xC7:
                        case 0xC8:
                        case 0xC9:
                        case 0xCA:
                        case 0xCB:
                        case 0xCD:
                        case 0xCE:
                        case 0xCF:
                            // SOFn (Start Of Frame) markers.
                            length = bh.ReadUshort();
                            bh.Seek(1);
                            Height = bh.ReadUshort();
                            Width = bh.ReadUshort();
                            bh.Seek(length - 7);
                            break;

                        default:
                            // Irrelevant variable-length packets.
                            length = bh.ReadUshort();
                            bh.Seek(length - 2);
                            break;
                    }
                }
            }
        }
    }
}