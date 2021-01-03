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
using System.Linq;
using System.Reflection;
using ExcelDnaUnpack.ResourceHelpers;
using SevenZip.Compression.LZMA;

namespace ExcelDnaUnpack
{
    public class Program
    {
        private static string _xllFile;
        private static bool _overwrite;
        private static string _outFolder;
        private static bool _showHelp;

        public static int Main(string[] args)
        {
            try
            {
                var appVersion = typeof(Program).Assembly.GetCustomAttributes(true)
                    .OfType<AssemblyInformationalVersionAttribute>().Single().InformationalVersion;

                var appTitle = $"Excel-DNA Unpack Tool, version {appVersion}";

                Console.WriteLine(appTitle, ConsoleColor.White);
                Console.WriteLine();

                var options = new OptionSet
                {
                    { "<>", v => _xllFile = v },
                    { "xllFile=", "The XLL file to be unpacked; e.g. MyAddIn-packed.xll", v => _xllFile = v },
                    { "outFolder=", "[Optional] The folder into which the extracted files will be written; defaults to `.\\unpacked`", v => _outFolder = v },
                    { "overwrite", "[Optional] Allow existing files of the same name to be overwritten; defaults to `false`", v => _overwrite = true },
                    { "help",  "Show this message and exit", v => _showHelp = !(v is null) },

                };

                options.Parse(args);

                if (args.Length < 1 || _showHelp)
                {
                    ShowHelp(options);
                    return 1;
                }

                Unpack();
                return 0;
            }
            catch (ApplicationException ex)
            {
                Console.WriteLine(ex.Message, ConsoleColor.Red);
                Console.WriteLine();
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}{Environment.NewLine}{ex}", ConsoleColor.Red);
                return 1;
            }
        }

        private static void Unpack()
        {
            if (string.IsNullOrWhiteSpace(_xllFile))
            {
                throw new ApplicationException("The XLL file is required.");
            }

            if (!File.Exists(_xllFile))
            {
                throw new ApplicationException($"The XLL file '{_xllFile}' doesn't exist.");
            }

            if (string.IsNullOrWhiteSpace(_outFolder))
            {
                _outFolder = Path.Combine(Path.GetFullPath(Environment.CurrentDirectory), "unpacked");
            }

            Console.Write($"Analyzing {Path.GetFileName(_xllFile)} . . . ");

            ExtractAllResources();
        }

        private static void ExtractAllResources()
        {
            var foundMainDnaResource = false;

            using (var reader = new ResourceReader(_xllFile))
            {
                Console.WriteLine("OK", ConsoleColor.Green);
                Console.WriteLine();

                if (reader.RecordCount > 0 && !Directory.Exists(_outFolder))
                {
                    Directory.CreateDirectory(_outFolder);
                }

                IResourceRecord resource;
                while ((resource = reader.Read()) != null)
                {
                    var outputFileName = GetFileNameWithExtension(resource);
                    Console.Write($"Extracting {Path.GetFileName(outputFileName)} ({resource.Type}) . . . ");

                    if ("__MAIN__.dna".Equals(outputFileName, StringComparison.Ordinal))
                    {
                        foundMainDnaResource = true;
                    }

                    var outputFilePath = Path.Combine(_outFolder, outputFileName);

                    if (File.Exists(outputFilePath) && !_overwrite)
                    {
                        Console.WriteLine("Error");
                        throw new ApplicationException(
                            $"The file {outputFileName} already exists and --overwrite was not specified");
                    }

                    using (var stream = new FileStream(outputFilePath, FileMode.Create))
                    {
                        var binaryData = resource.IsLzmaCompressed()
                            ? SevenZipHelper.Decompress(resource.BinaryData)
                            : resource.BinaryData;
                        stream.Write(binaryData, 0, binaryData.Length);
                    }

                    Console.WriteLine("OK", ConsoleColor.Green);
                }

                if (!foundMainDnaResource)
                {
                    Console.WriteLine($"__MAIN__.dna file not found. Are you sure {_xllFile} was built with Excel-DNA?",
                        ConsoleColor.Yellow);
                }
            }
        }

        private static void ShowHelp(OptionSet options)
        {
            Console.WriteLine("ExcelDna-Unpack is a command-line utility to extract the contents of Excel-DNA add-ins packed with ExcelDnaPack.");
            Console.WriteLine();
            Console.Write("Usage: ");
            Console.WriteLine("exceldna-unpack [<options>]", ConsoleColor.White);
            Console.WriteLine();
            Console.WriteLine("Where [<options>] is any of: ");
            Console.WriteLine();

            options.WriteOptionDescriptions(Console.Out);

            Console.WriteLine();
            Console.WriteLine(@"Example: exceldna-unpack --xllFile=MyAddIns\FirstAddIn-packed.xll");
            Console.WriteLine(@"         The extracted files will be saved to MyAddIns\unpacked");
        }

        private static string GetFileNameWithExtension(IResourceRecord resource)
        {
            var fileName = resource.Name;

            switch (resource.Type.ToUpperInvariant())
            {
                case "DNA":
                    {
                        fileName += ".dna";
                        break;
                    }

                case "ASSEMBLY":
                case "ASSEMBLY_LZMA":
                    {
                        fileName += ".dll";
                        break;
                    }

                case "TYPELIB":
                    {
                        fileName += ".tlb";
                        break;
                    }

                case "CONFIG":
                    {
                        fileName += ".config";
                        break;
                    }

                default:
                    {
                        var extension = Path.GetExtension(fileName);

                        if (!string.IsNullOrWhiteSpace(extension))
                        {
                            switch (extension.ToUpperInvariant())
                            {
                                case ".CS":
                                case ".VB":
                                case ".DNA":
                                case ".PNG":
                                case ".JPG":
                                case ".GIF":
                                    {
                                        fileName = Path.ChangeExtension(fileName, extension.ToLowerInvariant());
                                        break;
                                    }
                            }
                        }

                        break;
                    }
            }

            return fileName;
        }
    }
}
