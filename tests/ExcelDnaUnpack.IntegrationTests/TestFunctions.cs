using ExcelDna.Integration;

namespace ExcelDnaUnpack.IntegrationTests
{
    public class TestFunctions
    {
        [ExcelFunction(Description = "SayHello Excel-DNA Test function")]
        public static string SayHello(string name)
        {
            return "Test Hello " + name;
        }
    }
}
