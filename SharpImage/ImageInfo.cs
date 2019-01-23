using System.IO;

namespace SharpImage
{
    public class ImageInfo
    {
        public enum Type
        {
            Unknown = 0,
            Bmp,
            Gif,
            Jpeg,
            Png,
            Tiff
        }

        public static ImageInfoBase GetInfo(Type type, Stream stream)
        {
            ImageInfoBase info;

            switch (type)
            {
                case Type.Bmp:
                    info = new BmpInfo();
                    break;
                case Type.Gif:
                    info = new GifInfo();
                    break;
                case Type.Jpeg:
                    info = new JpegInfo();
                    break;
                case Type.Png:
                    info = new PngInfo();
                    break;
                case Type.Tiff:
                    info = new TiffInfo();
                    break;
                default:
                    return null;
            }

            info.Init(stream);
            return info;
        }

        public static Type GetType(string filename)
        {
            var extension = Path.GetExtension(filename);
            if (extension == null)
            {
                return Type.Unknown;
            }

            switch (extension.ToLower())
            {
                case ".bmp":
                    return Type.Bmp;
                case ".gif":
                    return Type.Gif;
                case ".jpg":
                case ".jpeg":
                    return Type.Jpeg;
                case ".png":
                    return Type.Png;
                case ".tif":
                case ".tiff":
                    return Type.Tiff;
            }

            return Type.Unknown;
        }
    }
}