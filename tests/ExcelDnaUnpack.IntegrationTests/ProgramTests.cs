using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Text;
using NUnit.Framework;

namespace ExcelDnaUnpack.IntegrationTests
{
    [TestFixture]
    public class ProgramTests
    {
        private string _originalFilesPath;
        private string _packedFilesPath;
        private string _testOutputPath;

        private const string PackedFile32Bit = "ExcelDnaUnpack.IntegrationTests-AddIn-packed.xll";
        private const string PackedFile64Bit = "ExcelDnaUnpack.IntegrationTests-AddIn64-packed.xll";

        [SetUp]
        public void SetUp()
        {
            var currentDirectory = Environment.CurrentDirectory;
            _originalFilesPath = currentDirectory;
            _packedFilesPath = currentDirectory;
            _testOutputPath = Path.Combine(currentDirectory, "TestOutput");

            if (Directory.Exists(_testOutputPath))
            {
                foreach (var fileName in Directory.GetFiles(_testOutputPath))
                {
                    File.Delete(fileName);
                }
            }
            else
            {
                Directory.CreateDirectory(_testOutputPath);
            }
        }

        [Test]
        [TestCase(PackedFile32Bit)]
        [TestCase(PackedFile64Bit)]
        public void App_unpacks_all_expected_files(string packedFileName)
        {
            // Arrange
            var filesExpected = new[]
            {
                "__MAIN__.dna",
                "__MAIN__.config",
                "EXCELDNAUNPACK.INTEGRATIONTESTS.dll",
                "EXCELDNA.LOADER.dll",
                "EXCELDNA.INTEGRATION.dll",
                "TESTIMAGE_0.png",
                "SOURCEFILETOBEPACKED_1.cs",
            };

            var args = new[] { "--xllFile=" + packedFileName, "--outFolder=" + _testOutputPath };

            // Act
            Program.Main(args);

            // Assert
            var outputFiles = Directory.GetFiles(_testOutputPath).Select(Path.GetFileName).ToArray();

            Assert.That(outputFiles.Length, Is.EqualTo(filesExpected.Length));
            CollectionAssert.AreEquivalent(filesExpected, outputFiles);
        }

        [Test]
        [TestCase(PackedFile32Bit)]
        [TestCase(PackedFile64Bit)]
        public void Unpacked_files_are_identical_to_original_ones_except_for_the_MAIN_dna_file(string packedFileName)
        {
            // Arrange
            var filesToCompare = new[]
            {
                new Tuple<string, string, bool>("ExcelDnaUnpack.IntegrationTests-AddIn.xll.config", "__MAIN__.config", false),
                new Tuple<string, string, bool>("ExcelDnaUnpack.IntegrationTests.dll", "EXCELDNAUNPACK.INTEGRATIONTESTS.dll", false),
                new Tuple<string, string, bool>("TestImage.png", "TESTIMAGE_0.png", false),
                new Tuple<string, string, bool>("SourceFileToBePacked.cs", "SOURCEFILETOBEPACKED_1.cs", true),
            };

            var args = new[] { "--xllFile=" + Path.Combine(_packedFilesPath, packedFileName), "--outFolder=" + _testOutputPath };

            // Act
            Program.Main(args);

            // Assert
            foreach (var file in filesToCompare)
            {
                var originalFileName = Path.Combine(_originalFilesPath, file.Item1);

                // When packing source files, ExcelDnaPack reads it as UTF8, so bytes needs to be compared accordingly
                var originalFileBytes = file.Item3 ? Encoding.UTF8.GetBytes(File.ReadAllText(originalFileName)) : File.ReadAllBytes(originalFileName);

                var unpackedFileName = Path.Combine(_testOutputPath, file.Item2);
                var unpackedFileBytes = File.ReadAllBytes(unpackedFileName);

                Assert.That(unpackedFileBytes.Length, Is.EqualTo(originalFileBytes.Length));
                Assert.IsTrue(StructuralComparisons.StructuralEqualityComparer.Equals(originalFileBytes, unpackedFileBytes), file.Item2);
            }
        }
    }
}
