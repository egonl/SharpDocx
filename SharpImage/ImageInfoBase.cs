using System.IO;

namespace SharpImage
{
    public abstract class ImageInfoBase
    {
        public bool IsValid { get; protected set; }

        public int Width { get; protected set; }
        public int Height { get; protected set; }

        public int DpiH { get; protected set; }
        public int DpiV { get; protected set; }

        public abstract void Init(Stream stream);

        protected int PixelsPerMeterToPixelsPerInch(int pixelsPerMeter)
        {
            return (int) ((pixelsPerMeter * 254L + 5000) / 10000);
        }
    }
}