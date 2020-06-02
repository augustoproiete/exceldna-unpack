using System;
using System.IO;

namespace ExcelDnaUnpack
{
    internal static class Console
    {
        public static void Write(string value)
        {
            System.Console.Write(value);
        }

        public static void WriteLine(string value)
        {
            System.Console.WriteLine(value);
        }

        public static void WriteLine(string value, ConsoleColor foregroundColor)
        {
            var previousForegroundColor = System.Console.ForegroundColor;

            try
            {
                System.Console.ForegroundColor = foregroundColor;
                System.Console.WriteLine(value);
            }
            finally
            {
                System.Console.ForegroundColor = previousForegroundColor;
            }
        }

        public static void WriteLine()
        {
            System.Console.WriteLine();
        }

        public static TextWriter Out => System.Console.Out;
    }
}
