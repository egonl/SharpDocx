using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace SharpImage
{
    public class ImageInfo
    {
        public static ImageInfoBase GetInfo(ImageType type, Stream stream)
        {
            ImageInfoBase info;

            switch (type)
            {
                case ImageType.Bmp:
                    info = new BmpInfo();
                    break;
                case ImageType.Gif:
                    info = new GifInfo();
                    break;
                case ImageType.Jpeg:
                    info = new JpegInfo();
                    break;
                case ImageType.Png:
                    info = new PngInfo();
                    break;
                case ImageType.Tiff:
                    info = new TiffInfo();
                    break;
                case ImageType.Emf:
                    info = new EmfInfo();
                    break;
                default:
                    return null;
            }

            info.Init(stream);
            return info;
        }

        public static ImageInfoBase GetInfo(Stream s)
        {
            ImageInfoBase info;

            if (!(info = GetInfo(s, new PngInfo())).IsValid)
                if (!(info = GetInfo(s, new BmpInfo())).IsValid)
                    if (!(info = GetInfo(s, new JpegInfo())).IsValid)
                        if (!(info = GetInfo(s, new GifInfo())).IsValid)
                            if (!(info = GetInfo(s, new TiffInfo())).IsValid)
                                if (!(info = GetInfo(s, new EmfInfo())).IsValid)
                                    return null;

            return info;
        }

        private static ImageInfoBase GetInfo(Stream stream, ImageInfoBase info)
        {
            info.Init(stream);
            return info;
        }

        public static ImageType GetType(string extension)
        {
            if (extension == null)
            {
                return ImageType.Unknown;
            }

            extension = extension.Replace(".", "").ToLower();
            switch (extension)
            {
                case "bmp":
                    return ImageType.Bmp;
                case "gif":
                    return ImageType.Gif;
                case "jpg":
                case "jpeg":
                    return ImageType.Jpeg;
                case "png":
                    return ImageType.Png;
                case "tif":
                case "tiff":
                    return ImageType.Tiff;
                case "emf":
                    return ImageType.Emf;
            }

            return ImageType.Unknown;
        }
    }
}
