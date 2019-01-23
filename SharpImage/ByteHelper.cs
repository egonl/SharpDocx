using System;
using System.IO;
using System.Text;

namespace SharpImage
{
    public class ByteHelper : IDisposable
    {
        public bool IsLsbf = true; // Least significat byte first / little endian / Intel format.

        public long Offset = 0L;

        public Stream Stream { get; set; }

        private readonly BinaryReader _binaryReader;

        public ByteHelper(Stream stream)
        {
            Stream = stream;
            _binaryReader = new BinaryReader(stream, Encoding.UTF8, true);
        }

        public void Dispose()
        {
            if (_binaryReader != null)
            {
                _binaryReader.Dispose();
            }
        }

        public void Seek(long offset, SeekOrigin origin = SeekOrigin.Current)
        {
            if (origin == SeekOrigin.Begin)
            {
                Stream.Seek(Offset + offset, origin);
            }
            else
            {
                Stream.Seek(offset, origin);
            }
        }

        public byte ReadByte()
        {
            return _binaryReader.ReadByte();
        }

        public byte[] ReadBytes(int length)
        {
            return _binaryReader.ReadBytes(length);
        }

        public ushort ReadUshort()
        {
            if (IsLsbf)
            {
                return _binaryReader.ReadUInt16();
            }

            var bytes = _binaryReader.ReadBytes(2);
            Array.Reverse(bytes);
            return BitConverter.ToUInt16(bytes, 0);
        }

        public uint ReadUint()
        {
            if (IsLsbf)
            {
                return _binaryReader.ReadUInt32();
            }

            var bytes = _binaryReader.ReadBytes(4);
            Array.Reverse(bytes);
            return BitConverter.ToUInt32(bytes, 0);
        }

        public string ReadAscii(int length)
        {
            var bytes = _binaryReader.ReadBytes(length);
            return Encoding.ASCII.GetString(bytes);
        }

        public static uint GetUintMsbf(byte[] bytes, int index = 0)
        {
            uint value0 = bytes[index];
            uint value1 = bytes[index + 1];
            uint value2 = bytes[index + 2];
            uint value3 = bytes[index + 3];

            value0 <<= 24;
            value1 <<= 16;
            value2 <<= 8;

            return value0 | value1 | value2 | value3;
        }

        public static ushort GetUshortMsbf(byte[] bytes, int index = 0)
        {
            ushort value0 = bytes[index];
            ushort value1 = bytes[index + 1];

            value0 <<= 8;

            return (ushort) (value0 | value1);
        }

        public static uint GetUintLsbf(byte[] bytes, int index = 0)
        {
            uint value0 = bytes[index + 3];
            uint value1 = bytes[index + 2];
            uint value2 = bytes[index + 1];
            uint value3 = bytes[index];

            value0 <<= 24;
            value1 <<= 16;
            value2 <<= 8;

            return value0 | value1 | value2 | value3;
        }

        public static ushort GetUshortLsbf(byte[] bytes, int index = 0)
        {
            ushort value0 = bytes[index + 1];
            ushort value1 = bytes[index];

            value0 <<= 8;

            return (ushort) (value0 | value1);
        }
    }
}