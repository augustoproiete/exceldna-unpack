#region Copyright 2014-2021 C. Augusto Proiete & Contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

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
