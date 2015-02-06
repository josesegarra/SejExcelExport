using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MySample
{
    public static class Extensions
    {

        // Simple extensions

        public static string Left(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        public static string Mid(this string value, int from)
        {
            if (string.IsNullOrEmpty(value)) return value;
            if (from >= value.Length) return "";
            return value.Substring(from);
        }

        public static string Mid(this string value, int from, int number)
        {
            if (string.IsNullOrEmpty(value)) return value;
            if (from >= value.Length) return "";
            if (number < 0) { from = from - number; number = -number; }
            if ((from + number) >= value.Length) return value.Substring(from);
            return value.Substring(from, number);
        }

    }
}
