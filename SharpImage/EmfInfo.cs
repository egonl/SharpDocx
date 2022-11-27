using System.IO;

namespace SharpImage
{
    public class EmfInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            // Inspired by https://www.codeproject.com/Articles/1307140/Parse-understand-and-demystify-Enhanced-Meta-Files

            using (var bh = new ByteHelper(stream))
            {
                bh.Seek(0, SeekOrigin.Begin);

                var emfHeader = bh.ReadUint();
                var headerSize = bh.ReadUint();
                IsValid = emfHeader == 1 && headerSize >= 80 && headerSize < 1024;

                if (!IsValid)
                {
                    return;
                }

                var boundsLeft = (int)bh.ReadUint();
                var boundsTop = (int)bh.ReadUint();
                var boundsRight = (int)bh.ReadUint();
                var boundsBottom = (int)bh.ReadUint();
                Width = boundsRight - boundsLeft + 2;
                Height = boundsBottom - boundsTop + 2;

                bh.Seek(48, SeekOrigin.Current);

                var deviceWidth = (int)bh.ReadUint();
                var deviceHeight = (int)bh.ReadUint();
                var mmWidth = (int)bh.ReadUint();
                int mmHeight = (int)bh.ReadUint();
                DpiH = (int)(((long)deviceWidth * 254 / mmWidth + 5) / 10);
                DpiV = (int)(((long)deviceHeight * 254 / mmHeight + 5) / 10);
            }
        }
    }
}
