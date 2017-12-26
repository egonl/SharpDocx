using System;

namespace Model
{
    public static class Extensions
    {
        public static string MyToString(this DateTime? dt)
        {
            return dt?.ToString("d") ?? string.Empty;
        }
    }
}