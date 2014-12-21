namespace ExcelDnaUnpack.ResourceHelpers
{
    internal interface IResourceRecord : IResourceHeader
    {
        byte[] BinaryData { get; set; }
    }
}