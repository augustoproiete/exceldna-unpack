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
                    .OfType<AssemblyFileVersionAttribute>().Single().Version;

                var appTitle = $"Excel-DNA Unpack Tool, version {appVersion}";

                Console.WriteLine(appTitle);
                Console.WriteLine();

                var options = new OptionSet
                {
                    { "<>", v => _xllFile = v },
                    { "xllFile=", "The XLL file to be unpacked; e.g. MyAddIn-packed.xll", v => _xllFile = v },
                    { "outFolder=", "[Optional] The folder into which the extracted files will be written; defaults to `.\\unpacked`", v => _outFolder = v },
                    { "overwrite", "[Optional] Allow existing files of the same name to be overwritten; default to `false`", v => _overwrite = true },
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
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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

            Console.Write("Analyzing {0} . . . ", Path.GetFileName(_xllFile));

            ExtractAllResources();
        }

        private static void ExtractAllResources()
        {
            using (var reader = new ResourceReader(_xllFile))
            {
                Console.WriteLine("OK");
                Console.WriteLine();

                if (reader.RecordCount > 0 && !Directory.Exists(_outFolder))
                {
                    Directory.CreateDirectory(_outFolder);
                }

                IResourceRecord resource;
                while ((resource = reader.Read()) != null)
                {
                    var outputFileName = GetFileNameWithExtension(resource);
                    Console.Write("Extracting {0} ({1}) . . . ", Path.GetFileName(outputFileName), resource.Type);

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

                    Console.WriteLine("OK");
                }
            }
        }

        private static void ShowHelp(OptionSet options)
        {
            Console.ResetColor();
            Console.WriteLine("ExcelDnaUnpack is a command-line utility to extract the contents of Excel-DNA add-ins packed with ExcelDnaPack.");
            Console.WriteLine();
            Console.Write("Usage: ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("ExcelDnaUnpack.exe [<options>]");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Where [<options>] is any of: ");
            Console.WriteLine();

            options.WriteOptionDescriptions(Console.Out);

            Console.WriteLine();
            Console.WriteLine(@"Example: ExcelDnaUnpack.exe --xllFile=MyAddins\FirstAddin-packed.xll");
            Console.WriteLine(@"         The extracted files will be saved to MyAddins\unpacked");
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
