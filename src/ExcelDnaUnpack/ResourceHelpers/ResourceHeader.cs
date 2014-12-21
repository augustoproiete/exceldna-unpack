using System.Diagnostics;

namespace ExcelDnaUnpack.ResourceHelpers
{
    [DebuggerDisplay("Name = {Name}, Type = {Type}")]
    internal class ResourceHeader : IResourceHeader
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }
}