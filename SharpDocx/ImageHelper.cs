using System;
using System.IO;
#if NET35 || NET45
using System.Windows.Media.Imaging;
#else
using SharpImage;
#endif
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
#if NET35 || NET45
            var img = new BitmapImage();

            using (imageStream)
            {
                img.BeginInit();
                img.StreamSource = imageStream;
                img.EndInit();

                imageStream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(imageStream);
                // imagePart will also dispose the stream.
            }

            var widthPx = img.PixelWidth;
            var heightPx = img.PixelHeight;
            var horzRezDpi = img.DpiX;
            var vertRezDpi = img.DpiY;
#else
            ImageInfoBase imageInfo = null;

            using (imageStream)
            {
                var type = GetImageInfoType(imagePartType);
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
#endif
            const int emusPerInch = 914400;
            var widthEmus = (long) (widthPx / horzRezDpi * emusPerInch);
            var heightEmus = (long) (heightPx / vertRezDpi * emusPerInch);

            if (widthEmus > maxWidthInEmus)
            {
                var ratio = heightEmus * 1.0m / widthEmus;
                widthEmus = maxWidthInEmus;
                heightEmus = (long) (widthEmus * ratio);
            }

            if (percentage != 100)
            {
                widthEmus = widthEmus * percentage / 100;
                heightEmus = heightEmus * percentage / 100;
            }

            var drawing = GetDrawing(package.MainDocumentPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
            return drawing;
        }

        public static ImagePartType GetImagePartType(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            if (extension == null)
            {
                throw new ArgumentException("Unknown extension", filePath);
            }

            // TODO: test which types actually work.
            ImagePartType? type = null;
            switch (extension.ToLower())
            {
                case ".bmp":
                    type = ImagePartType.Bmp;
                    break;
                case ".emf":
                    type = ImagePartType.Emf;
                    break;
                case ".gif":
                    type = ImagePartType.Gif;
                    break;
                case ".ico":
                case ".icon":
                    type = ImagePartType.Icon;
                    break;
                case ".jpg":
                case ".jpeg":
                    type = ImagePartType.Jpeg;
                    break;
                case ".pcx":
                    type = ImagePartType.Pcx;
                    break;
                case ".png":
                    type = ImagePartType.Png;
                    break;
                case ".tif":
                case ".tiff":
                    type = ImagePartType.Tiff;
                    break;
                case ".wmf":
                    type = ImagePartType.Wmf;
                    break;
            }

            if (type == null)
            {
                throw new ArgumentException("Unknown extension", filePath);
            }

            return type.Value;
        }

#if !(NET35 || NET45)
        public static ImageInfo.Type GetImageInfoType(ImagePartType imagePartType)
        {
            switch (imagePartType)
            {
                case ImagePartType.Bmp:
                    return ImageInfo.Type.Bmp;

                case ImagePartType.Gif:
                    return ImageInfo.Type.Gif;

                case ImagePartType.Jpeg:
                    return ImageInfo.Type.Jpeg;

                case ImagePartType.Png:
                    return ImageInfo.Type.Png;

                case ImagePartType.Tiff:
                    return ImageInfo.Type.Tiff;
            }

            return ImageInfo.Type.Unknown;
        }
#endif

        private static Drawing GetDrawing(string relationshipId, long widthEmus, long heightEmus)
        {
            return new Drawing(
                new DW.Inline(
                    new DW.Extent {Cx = widthEmus, Cy = heightEmus},
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
                        new A.GraphicFrameLocks {NoChangeAspect = true}),
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
                                                new A.BlipExtension {Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"}))
                                        {
                                            Embed = relationshipId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset {X = 0, Y = 0},
                                            new A.Extents {Cx = widthEmus, Cy = heightEmus}),
                                        new A.PresetGeometry(
                                                new A.AdjustValueList())
                                            {Preset = A.ShapeTypeValues.Rectangle}))
                            )
                            {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
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