using System;
using System.Collections.Generic;

namespace ExcelDnaUnpack.ResourceHelpers
{
    internal interface IResourceReader : IEnumerable<IResourceRecord>, IDisposable
    {
        IResourceRecord Read();
        bool CanRead();
        void Close();

        int RecordCount { get; }
    }
}