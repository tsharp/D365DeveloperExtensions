using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

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

        public static string WriteTempFile(string name, byte[] content)
        {
            try
            {
                var tempFolder = Path.GetTempPath();
                string fileName = Path.GetFileName(name);
                if (string.IsNullOrEmpty(fileName))
                    fileName = Guid.NewGuid().ToString();
                var tempFile = Path.Combine(tempFolder, fileName);
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
                File.WriteAllBytes(tempFile, content);

                return tempFile;
            }
            catch (Exception)
            {
                MessageBox.Show("Error writing temp file");
                throw;
            }
        }

        public static void WriteFileToDisk(string path, byte[] content)
        {
            try
            {
                File.WriteAllBytes(path, content);
            }
            catch (Exception)
            {
                MessageBox.Show("Error writing file");
                throw;
            }
        }
    }
}
