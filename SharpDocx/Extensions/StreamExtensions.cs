#if NET35
using System.IO;

namespace SharpDocx.Extensions
{
    public static class StreamExtensions
    {
        public static void CopyTo(this Stream source, Stream destination)
        {
            byte[] buffer = new byte[4096];
            int bytesRead;

            while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                destination.Write(buffer, 0, bytesRead);
            }
        }
    }
}
#endif