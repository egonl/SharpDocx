using System.IO;

namespace SharpImage
{
    public class GifInfo : ImageInfoBase
    {
        public override void Init(Stream stream)
        {
            using (var bh = new ByteHelper(stream))
            {
                bh.Seek(0, SeekOrigin.Begin);

                var imageHeader = bh.ReadBytes(6);

                // Should be GIF87a or GIF89a.
                IsValid = imageHeader[0] == 'G' && imageHeader[1] == 'I' && imageHeader[2] == 'F' && 
                          imageHeader[3] == '8' && (imageHeader[4] == '7' || imageHeader[4] == '9') && 
                          imageHeader[5] == 'a';

                if (!IsValid)
                {
                    return;
                }

                Width = bh.ReadUshort();
                Height = bh.ReadUshort();

                // GIF files don't specify a resolution.
                DpiH = 72;
                DpiV = 72;
            }
        }
    }
}