using System.IO;

namespace SharpImage
{
    public abstract class ImageInfoBase
    {
        public ImageType Type { get; protected set; }

        public bool IsValid { get; protected set; }

        public int Width { get; protected set; }
        public int Height { get; protected set; }

        public int DpiH { get; protected set; } = 96;
        public int DpiV { get; protected set; } = 96;

        public abstract void Init(Stream stream);

        protected int PixelsPerMeterToPixelsPerInch(int pixelsPerMeter)
        {
            return (int) ((pixelsPerMeter * 254L + 5000) / 10000);
        }
    }
}