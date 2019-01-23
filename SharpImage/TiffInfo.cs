using System;
using System.IO;

namespace SharpImage
{
    public class TiffInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            ushort resolutionUnit = 0;

            using (var bh = new ByteHelper(stream))
            {
                bh.Seek(0, SeekOrigin.Begin);

                var byteOrder = bh.ReadAscii(2);
                if (byteOrder != "II" && byteOrder != "MM")
                {
                    return;
                }

                bh.IsLsbf = byteOrder == "II";

                bh.Seek(2);

                var ifdOffset = bh.ReadUint();
                bh.Seek(ifdOffset, SeekOrigin.Begin);

                var entryCount = bh.ReadUshort();

                for (ushort i = 0; i < entryCount; ++i)
                {
                    var entryTag = bh.ReadUshort();
                    var fieldType = bh.ReadUshort();
                    var numberOfComponents = bh.ReadUint();

                    switch (entryTag)
                    {
                        case (int) Tags.ImageWidth:
                            switch (fieldType)
                            {
                                case (ushort) FieldTypes.Short:
                                    Width = bh.ReadUshort();
                                    bh.Seek(2);
                                    break;

                                case (ushort) FieldTypes.Long:
                                    Width = (int) bh.ReadUint();
                                    break;
                            }

                            break;

                        case (int) Tags.ImageHeight:
                            switch (fieldType)
                            {
                                case (ushort) FieldTypes.Short:
                                    Height = bh.ReadUshort();
                                    bh.Seek(2);
                                    break;

                                case (ushort) FieldTypes.Long:
                                    Height = (int) bh.ReadUint();
                                    break;
                            }

                            break;

                        case (int) Tags.XResolution:
                        {
                            // Field type is always rational.
                            var tagDataOffset = bh.ReadUint();
                            var currentOffset = stream.Position;
                            stream.Seek(tagDataOffset, SeekOrigin.Begin);
                            var numerator = bh.ReadUint();
                            var denominator = bh.ReadUint();
                            DpiH = (int) Math.Round((double) numerator / denominator);
                            stream.Seek(currentOffset, SeekOrigin.Begin);
                            break;
                        }

                        case (int) Tags.YResolution:
                        {
                            // Field type is always rational.
                            var tagDataOffset = bh.ReadUint();
                            var currentOffset = stream.Position;
                            stream.Seek(tagDataOffset, SeekOrigin.Begin);
                            var numerator = bh.ReadUint();
                            var denominator = bh.ReadUint();
                            DpiV = (int) Math.Round((double) numerator / denominator);
                            stream.Seek(currentOffset, SeekOrigin.Begin);
                            break;
                        }

                        case (int) Tags.ResolutionUnit:
                            resolutionUnit = bh.ReadUshort();
                            bh.Seek(2);
                            break;

                        default:
                            bh.Seek(4);
                            break;
                    }
                }
            }

            // https://www.awaresystems.be/imaging/tiff/tifftags/resolutionunit.html
            if (resolutionUnit == 0 || resolutionUnit == 1)
            {
                DpiH = DpiV = 72;
            }
            else if (resolutionUnit == 3)
            {
                DpiH = PixelsPerMeterToPixelsPerInch(DpiH * 100);
                DpiV = PixelsPerMeterToPixelsPerInch(DpiV * 100);
            }

            IsValid = true;
        }

        private enum FieldTypes
        {
            Short = 3,
            Long = 4,
            Rational = 5
        }

        private enum Tags
        {
            ImageWidth = 256,
            ImageHeight = 257,
            XResolution = 282,
            YResolution = 283,
            ResolutionUnit = 296
        }
    }
}