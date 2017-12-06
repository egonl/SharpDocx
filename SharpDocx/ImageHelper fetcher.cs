using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace JwpMailFetcher.Imap
{
    public static class ImageHelper
    {
        #region public methods
        
        public static bool IsBmp(byte[] image)
        {
            try
            {
                return image[0] == 'B' && image[1] == 'M';
            }
            catch
            {
                return false;
            }
        }

        public static bool IsGif(byte[] image)
        {
            try
            {
                // GIF87a/GIF89a file
                return 
                    image[0] == 'G' && image[1] == 'I' && image[2] == 'F' && 
                    image[3] == '8' && (image[4] == '7' || image[4] == '9')  && image[5] == 'a';
            }
            catch
            {
                return false;
            }
        }

        public static bool IsPng(byte[] image)
        {
            try
            {
                return 
                    image[1] == 'P' && image[2] == 'N' && image[3] == 'G' &&
                    image[12] == 'I' && image[13] == 'H' && image[14] == 'D' && image[15] == 'R';
            }
            catch
            {
                return false;
            }
        }

        public static bool IsJpeg(byte[] image)
        {
            try
            {
                // Start of Image (SOI) marker (FFD8).
                ushort soi = Get2BytesValueMsbf(image, 0);
                if (soi != 0xFFD8) return false;

                // JFIF marker (FFE0), EXIF marker (FFE1) of ...
                ushort marker = Get2BytesValueMsbf(image, 2);
                if ((marker & 0xFFE0) == 0xFFE0) return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        public static Size GetBmpDimensions(byte[] image)
        {
            if (!IsBmp(image))
                throw new ApplicationException("Unexpected value in bitmap file");

            ushort width = Get2BytesValueLsbf(image, 18);
            ushort height = Get2BytesValueLsbf(image, 22);
            return new Size(width, height);
        }

        public static Size GetGifDimensions(byte[] image)
        {
            if (!IsGif(image))
                throw new ApplicationException("Unexpected value in Gif file");

            ushort width = Get2BytesValueLsbf(image, 6);
            ushort height = Get2BytesValueLsbf(image, 8);
            return new Size(width, height);
        }

        public static Size GetJpegDimensions(byte[] image)
        {
            int index = 0;
            // keep reading packets until we find one that contains Size info
            for (; ; )
            {
                byte code = image[index];
                index++;

                if (code != 0xFF) throw new ApplicationException("Unexpected value in Jpeg file");
                code = image[index];
                index++;

                switch (code)
                {
                    // filler byte
                    case 0xFF:
                        index--;
                        break;
                    // packets without data
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
                        break;
                    // packets with size information
                    case 0xC0:
                    case 0xC1:
                    case 0xC2:
                    case 0xC3:
                    case 0xC4:
                    case 0xC5:
                    case 0xC6:
                    case 0xC7:
                    case 0xC8:
                    case 0xC9:
                    case 0xCA:
                    case 0xCB:
                    case 0xCC:
                    case 0xCD:
                    case 0xCE:
                    case 0xCF:
                        ushort h = Get2BytesValueMsbf(image, index + 3);
                        ushort w = Get2BytesValueMsbf(image, index + 5);
                        return new Size(w, h);
                    // irrelevant variable-length packets
                    default:
                        int len = Get2BytesValueMsbf(image, index);
                        index += len;
                        break;
                }
            }
        }

        public static Size GetPngDimensions(byte[] image)
        {
            if (!IsPng(image))
                throw new ApplicationException("Unexpected value in Png file");
           
            uint width = Get4BytesValueMsbf(image, 16);
            uint height = Get4BytesValueMsbf(image, 20);

            return new Size((int)width, (int)height);
        }

        public static Size GetTiffDimensions(byte[] image)
        {
            //The byte order used within the file. Legal values are:
            //“II” (4949.H) 
            //“MM” (4D4D.H)
            //In the “II” format, byte order is always from the least significant byte to the most
            //significant byte, for both 16-bit and 32-bit integers This is called little-endian byte
            //order. In the “MM” format, byte order is always from most significant to least
            //significant, for both 16-bit and 32-bit integers. This is called big-endian byte
            //order.
            bool isMsbf = (image[0].Equals(77) && image[1].Equals(77)); 

            if (!Get2BytesValue(image, 2, isMsbf).Equals(42))
                throw new ApplicationException("Unexpected value in Tiff file");

            int width = -1;
            int height = -1;

            uint offset = Get4BytesValue(image, 4, isMsbf);
            
            // Read Image File Directories (IFD)
            while (!offset.Equals(0) && (width < 0 || height < 0))
            {
                if (offset > image.Length)
                    throw new ApplicationException("Tiff file is corrupt"); // So it will not be an endless loop if the file content is corrupt

                ushort nrOfDirectories = Get2BytesValue(image, (int)offset, isMsbf);
                offset += 2;

                for (uint i = 0; i < nrOfDirectories && (width < 0 || height < 0); i++)
                {
                    ushort ifdEntryTag = Get2BytesValue(image, (int)offset, isMsbf);
                    ushort fieldType = Get2BytesValue(image, (int)offset + 2, isMsbf);

                    switch (ifdEntryTag)
                    {

                        case 256: // ImageWidth
                            switch (fieldType)
                            {
                                case 3: // Short
                                    width = Get2BytesValue(image, (int)offset + 8, isMsbf);
                                    break;

                                case 4: // Long
                                    width = (int)Get4BytesValue(image, (int)offset + 8, isMsbf);
                                    break;
                            }
                            break;

                        case 257: // ImageHeight
                            switch (fieldType)
                            {
                                case 3: // Short
                                    height = Get2BytesValue(image, (int)offset + 8, isMsbf);
                                    break;

                                case 4: // Long
                                    height = (int)Get4BytesValue(image, (int)offset + 8, isMsbf);
                                    break;
                            }
                            break;
                    }
                    offset += 12;
                }
                // Get the offset to the next set of IFDs
                offset = Get4BytesValue(image, (int)offset, isMsbf);
            }

            if (height < 0)
                throw new ApplicationException("No height found in Tiff file");

            if (width < 0)
                throw new ApplicationException("No width found in Tiff file");

            return new Size(width, height);
        }

        #endregion


        #region private methods

        private static uint Get4BytesValue(byte[] image, int index, bool isMsbf)
        {
            if (isMsbf)
                return Get4BytesValueMsbf(image, index);
            else
                return Get4BytesValueLsbf(image, index);
        }


        private static ushort Get2BytesValue(byte[] image, int index, bool isMsbf)
        {
            if (isMsbf)
                return Get2BytesValueMsbf(image, index);
            else
                return Get2BytesValueLsbf(image, index);
        }


        /// <summary>
        /// Get a unsigned integer value of 4 bytes (Most significant byte first) 
        /// </summary>
        /// <param name="image"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private static uint Get4BytesValueMsbf(byte[] image, int index)
        {
            uint value0 = image[index];
            uint value1 = image[index + 1];
            uint value2 = image[index + 2];
            uint value3 = image[index + 3];

            value0 <<= 24;
            value1 <<= 16;
            value2 <<= 8;

            return (uint)(value0 | value1 | value2 | value3);
        }


        /// <summary>
        /// Get a unsigned short integer value of 2 bytes (Most significant byte first)
        /// </summary>
        /// <param name="image"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private static ushort Get2BytesValueMsbf(byte[] image, int index)
        {
            ushort value0 = image[index];
            ushort value1 = image[index + 1];

            value0 <<= 8;

            return (ushort)(value0 | value1);
        }


        /// <summary>
        /// Get a unsigned integer value of 4 bytes (Least significant byte first)
        /// </summary>
        /// <param name="image"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private static uint Get4BytesValueLsbf(byte[] image, int index)
        {
            uint value0 = image[index + 3];
            uint value1 = image[index + 2];
            uint value2 = image[index + 1];
            uint value3 = image[index];

            value0 <<= 24;
            value1 <<= 16;
            value2 <<= 8;

            return (uint)(value0 | value1 | value2 | value3);
        }

        /// <summary>
        /// Get a unsigned short integer value of 2 bytes (Least significant byte first)
        /// </summary>
        /// <param name="image"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private static ushort Get2BytesValueLsbf(byte[] image, int index)
        {
            ushort value0 = image[index + 1];
            ushort value1 = image[index];

            value0 <<= 8;

            return (ushort)(value0 | value1);
        }

        #endregion
    }
}
