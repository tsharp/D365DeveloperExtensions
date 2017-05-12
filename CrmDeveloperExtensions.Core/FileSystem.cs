using System.Collections.Generic;
using System.IO;

namespace CrmDeveloperExtensions.Core
{
    public static class FileSystem
    {
        public static bool IsDirectoryEmpty(string path)
        {
            IEnumerable<string> items = Directory.EnumerateFileSystemEntries(path);
            using (IEnumerator<string> en = items.GetEnumerator())
                return !en.MoveNext();
        }
    }
}
