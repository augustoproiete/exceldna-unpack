namespace ExcelDnaUnpack.ResourceHelpers
{
    internal class ResourceRecord : ResourceHeader, IResourceRecord
    {
        public byte[] BinaryData { get; set; }
    }
}