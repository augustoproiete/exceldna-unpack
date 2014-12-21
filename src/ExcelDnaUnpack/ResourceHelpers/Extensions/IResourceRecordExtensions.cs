using System;

// ReSharper disable once CheckNamespace
namespace ExcelDnaUnpack.ResourceHelpers
{
    // ReSharper disable once InconsistentNaming
    internal static class IResourceRecordExtensions
    {
        public static bool IsLzmaCompressed(this IResourceRecord resource)
        {
            return resource.Type.EndsWith("_LZMA", StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
