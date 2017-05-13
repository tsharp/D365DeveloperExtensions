using System;
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

        public static DirectoryInfo GetDirectory(string input)
        {
            DirectoryInfo directory = new DirectoryInfo(input);
            if (!directory.Exists)
                throw new Exception("Unable to get directory from string");

            return directory;
        }
    }
}
