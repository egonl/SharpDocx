using System;
using System.IO;
using SharpImage;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace SharpDocx
{
    public static class ImageHelper
    {
        public static Drawing CreateDrawing(
            WordprocessingDocument package,
            Stream imageStream,
            ImagePartType imagePartType,
            int percentage, long maxWidthInEmus)
        {
            var imagePart = package.MainDocumentPart.AddImagePart(imagePartType);

            ImageInfoBase imageInfo = null;

            using (imageStream)
            {
                var type = GetImageType(imagePartType);
                imageInfo = ImageInfo.GetInfo(type, imageStream);
                if (imageInfo == null)
                {
                    throw new ArgumentException("Unsupported image format.");
                }

                imageStream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(imageStream);
                // imagePart will also dispose the stream.
            }

            var widthPx = imageInfo.Width;
            var heightPx = imageInfo.Height;
            var horzRezDpi = imageInfo.DpiH;
            var vertRezDpi = imageInfo.DpiV;

            const int emusPerInch = 914400;
            var widthEmus = (long)widthPx * emusPerInch / horzRezDpi;
            var heightEmus = (long)heightPx * emusPerInch / vertRezDpi;

            if (widthEmus > maxWidthInEmus)
            {
                var ratio = heightEmus * 1.0m / widthEmus;
                widthEmus = maxWidthInEmus;
                heightEmus = (long)(widthEmus * ratio);
            }

            if (percentage != 100)
            {
                widthEmus = widthEmus * percentage / 100;
                heightEmus = heightEmus * percentage / 100;
            }

            var drawing = GetDrawing(package.MainDocumentPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
            return drawing;
        }

        internal static ImagePartType GetImagePartType(string extension)
        {
            if (extension == null)
            {
                throw new ArgumentException("Unknown extension", nameof(extension));
            }

            ImagePartType? type = null;

            extension = extension.Replace(".", "").ToLower();
            switch (extension)
            {
                case "bmp":
                    type = ImagePartType.Bmp;
                    break;
                case "emf": // untested
                    type = ImagePartType.Emf;
                    break;
                case "gif":
                    type = ImagePartType.Gif;
                    break;
                case "ico": // untested
                case "icon":
                    type = ImagePartType.Icon;
                    break;
                case "jpg":
                case "jpeg":
                    type = ImagePartType.Jpeg;
                    break;
                case "pcx": // untested
                    type = ImagePartType.Pcx;
                    break;
                case "png":
                    type = ImagePartType.Png;
                    break;
                case "tif":
                case "tiff":
                    type = ImagePartType.Tiff;
                    break;
                case "wmf":
                    type = ImagePartType.Wmf;
                    break;
            }

            if (type == null)
            {
                throw new ArgumentException("Unknown extension", nameof(extension));
            }

            return type.Value;
        }

        internal static ImageType GetImageType(ImagePartType imagePartType)
        {
            switch (imagePartType)
            {
                case ImagePartType.Bmp:
                    return ImageType.Bmp;

                case ImagePartType.Gif:
                    return ImageType.Gif;

                case ImagePartType.Jpeg:
                    return ImageType.Jpeg;

                case ImagePartType.Png:
                    return ImageType.Png;

                case ImagePartType.Tiff:
                    return ImageType.Tiff;

                case ImagePartType.Emf:
                    return ImageType.Emf;
            }

            return ImageType.Unknown;
        }

        internal static ImagePartType GetImagePartType(ImageType imageType)
        {
            switch (imageType)
            {
                case ImageType.Bmp:
                    return ImagePartType.Bmp;

                case ImageType.Gif:
                    return ImagePartType.Gif;

                case ImageType.Jpeg:
                    return ImagePartType.Jpeg;

                case ImageType.Png:
                    return ImagePartType.Png;

                case ImageType.Tiff:
                    return ImagePartType.Tiff;

                case ImageType.Emf:
                    return ImagePartType.Emf;
            }

            return (ImagePartType) (-1);
        }

        private static Drawing GetDrawing(string relationshipId, long widthEmus, long heightEmus)
        {
            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmus, Cy = heightEmus },
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = 1,
                        Name = "Picture 1"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties
                                        {
                                            Id = 0,
                                            Name = "New Bitmap Image.png"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                        {
                                            Embed = relationshipId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset { X = 0, Y = 0 },
                                            new A.Extents { Cx = widthEmus, Cy = heightEmus }),
                                        new A.PresetGeometry(
                                                new A.AdjustValueList())
                                            {Preset = A.ShapeTypeValues.Rectangle}))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = 0,
                    DistanceFromBottom = 0,
                    DistanceFromLeft = 0,
                    DistanceFromRight = 0,
                    EditId = "50D07946"
                });
        }
    }
}