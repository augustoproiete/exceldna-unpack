using System;
using System.IO;
using System.Linq;
using Xunit;
using FluentAssertions;

namespace ExcelDnaUnpack.Tests
{
    public class ExcelDnaUnpackTests
    {
        private readonly string _originalFilesPath;
        private readonly string _packedFilesPath;
        private readonly string _testOutputPath;

        private const string PackedFile32Bit = "ExcelDnaUnpack.Tests.ExcelAddIn-packed.xll";
        private const string PackedFile64Bit = "ExcelDnaUnpack.Tests.ExcelAddIn64-packed.xll";

        public ExcelDnaUnpackTests()
        {
            var currentDirectory = Environment.CurrentDirectory;
            _originalFilesPath = currentDirectory;
            _packedFilesPath = currentDirectory;
            _testOutputPath = Path.Combine(currentDirectory, "TestOutput");

            if (Directory.Exists(_testOutputPath))
            {
                foreach (var fileName in Directory.EnumerateFiles(_testOutputPath))
                {
                    File.Delete(fileName);
                }
            }
            else
            {
                Directory.CreateDirectory(_testOutputPath);
            }
        }

        [InlineData(PackedFile32Bit)]
        [InlineData(PackedFile64Bit)]
        [Theory]
        public void App_unpacks_all_expected_files(string packedFileName)
        {
            var filesExpected = new[]
            {
                "__MAIN__.dna",
                "__MAIN__.config",
                "EXCELDNAUNPACK.TESTS.EXCELADDIN.dll",
                "EXCELDNA.LOADER.dll",
                "EXCELDNA.INTEGRATION.dll",
                "TESTIMAGE_0.png",
                "SOURCEFILETOBEPACKED_1.cs",
            };

            var args = new[]
            {
                $"--xllFile={packedFileName}",
                $"--outFolder={_testOutputPath}",
            };

            var exitCode = Program.Main(args);
            exitCode.Should().Be(0);

            var outputFiles = Directory.EnumerateFiles(_testOutputPath)
                .Select(Path.GetFileName)
                .ToArray();

            outputFiles.Should().BeEquivalentTo(filesExpected);
        }

        [InlineData(PackedFile32Bit)]
        [InlineData(PackedFile64Bit)]
        [Theory]
        public void Unpacked_files_are_identical_to_original_ones_except_for_the_MAIN_dna_file(string packedFileName)
        {
            var filesToCompare = new[]
            {
                new Tuple<string, string>("ExcelDnaUnpack.Tests.ExcelAddIn.xll.config", "__MAIN__.config"),
                new Tuple<string, string>("ExcelDnaUnpack.Tests.ExcelAddIn.dll", "EXCELDNAUNPACK.TESTS.EXCELADDIN.dll"),
                new Tuple<string, string>("TestImage.png", "TESTIMAGE_0.png"),
                new Tuple<string, string>("SourceFileToBePacked.cs", "SOURCEFILETOBEPACKED_1.cs"),
            };

            var args = new[]
            {
                $"--xllFile={Path.Combine(_packedFilesPath, packedFileName)}",
                $"--outFolder={_testOutputPath}",
            };

            var exitCode = Program.Main(args);
            exitCode.Should().Be(0);

            foreach (var (originalFileName, fileNameInsidePacked) in filesToCompare)
            {
                var originalFilePath = Path.Combine(_originalFilesPath, originalFileName);
                var originalFileBytes = File.ReadAllBytes(originalFilePath);

                var unpackedFileName = Path.Combine(_testOutputPath, fileNameInsidePacked);
                var unpackedFileBytes = File.ReadAllBytes(unpackedFileName);

                unpackedFileBytes.Should().BeEquivalentTo(originalFileBytes, options => options.WithStrictOrdering(),
                    because: $"{fileNameInsidePacked} should be identical to {originalFileName}");
            }
        }
    }
}
