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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace ExcelDnaUnpack.ResourceHelpers
{
    internal class ResourceReader : IResourceReader
    {
        private IntPtr _fileHandle;
        private IList<ResourceHeader> _resourceHeaders;

        private int _enumerator;


        public ResourceReader(string fileName)
        {
            _fileHandle = LoadLibraryEx(fileName, IntPtr.Zero, ((uint)(LoadLibraryFlags.LoadLibraryAsDatafile | LoadLibraryFlags.LoadLibraryAsImageResource)));
            if (_fileHandle == IntPtr.Zero)
            {
                throw new Win32Exception("Unable to load file " + fileName);
            }

            ReadResourceHeaders();
        }

        ~ResourceReader()
        {
            Dispose(false);
        }

        public IResourceRecord Read()
        {
            if (!CanRead())
            {
                return null;
            }

            var resource = new ResourceRecord
            {
                Name = _resourceHeaders[_enumerator].Name,
                Type = _resourceHeaders[_enumerator].Type,
            };

            resource.BinaryData = ReadBinaryData(resource.Name, resource.Type);

            _enumerator++;

            return resource;
        }

        public bool CanRead()
        {
            return _enumerator < _resourceHeaders.Count;
        }

        public IEnumerator<IResourceRecord> GetEnumerator()
        {
            IResourceRecord record;
            while ((record = Read()) != null)
            {
                yield return record;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Close()
        {
            Dispose(true);
        }

        public int RecordCount => _resourceHeaders.Count;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void ReadResourceHeaders()
        {
            _resourceHeaders = new List<ResourceHeader>();

            var resourceTypes = GetResourceTypes();
            foreach (var type in resourceTypes)
            {
                var resourceNames = GetResourceNames(type);
                foreach (var name in resourceNames)
                {
                    var resourceHeader = new ResourceHeader
                    {
                        Name = name,
                        Type = type,
                    };

                    _resourceHeaders.Add(resourceHeader);
                }
            }
        }

        private IEnumerable<string> GetResourceNames(string resourceType)
        {
            var resourceNames = new List<string>();

            EnumResourceNamesWithName(_fileHandle, resourceType, (module, type, name, param) =>
            {
                var resourceName = GetResourceName(name);
                resourceNames.Add(resourceName);

                return true;
            }, IntPtr.Zero);

            return resourceNames;
        }

        private IEnumerable<string> GetResourceTypes()
        {
            var resourceTypes = new List<string>();
            EnumResourceTypes(_fileHandle, (module, type, param) =>
            {
                resourceTypes.Add(type);
                return true;
            }, IntPtr.Zero);

            return resourceTypes;
        }

        private static string GetResourceName(IntPtr value)
        {
            if (IsIntResource(value))
            {
                return $"#{value}";
            }

            return Marshal.PtrToStringUni(value);
        }

        private static bool IsIntResource(IntPtr value)
        {
            return (((ulong)value >> 16) == 0);
        }

        private byte[] ReadBinaryData(string name, string type)
        {
            var resourceHandleFind = FindResource(_fileHandle, name, type);
            if (resourceHandleFind == IntPtr.Zero)
            {
                var message =
                    $"FindResource couldn't find resource '{name}' of type '{type}'. GetLastError={Marshal.GetLastWin32Error()}";

                throw new Win32Exception(message);
            }

            var resourceHandleLoad = LoadResource(_fileHandle, resourceHandleFind);
            if (resourceHandleLoad == IntPtr.Zero)
            {
                var message =
                    $"LoadResource couldn't load resource '{name}' of type '{type}'. GetLastError={Marshal.GetLastWin32Error()}";

                throw new Win32Exception(message);
            }

            var resourceHandleLock = LockResource(resourceHandleLoad);
            if (resourceHandleLock == IntPtr.Zero)
            {
                var message =
                    $"LockResource couldn't lock resource '{name}' of type '{type}'. GetLastError={Marshal.GetLastWin32Error()}";

                throw new Win32Exception(message);
            }

            var resourceSize = SizeOfResource(_fileHandle, resourceHandleFind);
            var resourceBytes = new byte[resourceSize];

            Marshal.Copy(resourceHandleLock, resourceBytes, 0, (int)resourceSize);

            return resourceBytes;
        }

        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if (_fileHandle != IntPtr.Zero)
                {
                    FreeLibrary(_fileHandle);
                    _fileHandle = IntPtr.Zero;
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch
            {
                // Do nothing
            }
        }

        [DllImport("kernel32.dll", EntryPoint = "LoadLibraryEx", SetLastError = true)]
        private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", EntryPoint = "FreeLibrary", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", EntryPoint = "EnumResourceTypes", SetLastError = true)]
        private static extern bool EnumResourceTypes(IntPtr hModule, EnumResTypeProc lpEnumFunc, IntPtr lParam);

        [DllImport("kernel32.dll", EntryPoint = "EnumResourceNamesW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool EnumResourceNamesWithName(IntPtr hModule, string lpszType, EnumResNameProc lpEnumFunc, IntPtr lParam);

        [DllImport("Kernel32.dll", EntryPoint = "FindResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindResource(IntPtr hModule, string pName, string pType);
        
        [DllImport("Kernel32.dll", EntryPoint = "SizeofResource", SetLastError = true)]
        private static extern uint SizeOfResource(IntPtr hModule, IntPtr hResource);

        [DllImport("Kernel32.dll", EntryPoint = "LoadResource", SetLastError = true)]
        private static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResource);

        [DllImport("Kernel32.dll", EntryPoint = "LockResource")]
        private static extern IntPtr LockResource(IntPtr hGlobal);

        private delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);

        private delegate bool EnumResTypeProc(IntPtr hModule, string lpszType, IntPtr lParam);

        [Flags]
        private enum LoadLibraryFlags : uint
        {
            LoadLibraryAsDatafile = 0x00000002,
            LoadLibraryAsImageResource = 0x00000020,
        }
    }
}
