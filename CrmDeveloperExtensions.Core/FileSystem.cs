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
            string path = Path.GetDirectoryName(input);
            if (path == null)
                throw new Exception("Unable to get directory from string");

            DirectoryInfo directory = new DirectoryInfo(path);
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
                if (String.IsNullOrEmpty(fileName))
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

        public static string BoundFileToLocalPath(string boundFile, string projectPath)
        {
            string path = Path.GetDirectoryName(projectPath);
            if (path == null)
                return null;

            if (boundFile.StartsWith("/"))
                boundFile = boundFile.Substring(1);

            return Path.Combine(path, boundFile.Replace("/", "\\"));
        }

        public static string LocalPathToCrmPath(string projectPath, string filename)
        {
            return filename.Replace(projectPath, String.Empty).Replace("\\", "/");
        }
    }
}
