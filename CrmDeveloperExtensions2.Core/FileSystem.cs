using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace CrmDeveloperExtensions2.Core
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

        public static bool DoesFileExist(string[] files, bool checkAll)
        {
            foreach (string file in files)
            {
                bool exists = File.Exists(file);
                if (exists && !checkAll)
                    return true;

                if (!exists && checkAll)
                    return false;
            }

            return checkAll;
        }

        public static bool ConfirmOverwrite(string[] files, bool checkAll)
        {
            List<string> existingFiles = new List<string>();

            foreach (string file in files)
            {
                bool exists = DoesFileExist(new[] { file }, true);
                if (exists)
                    existingFiles.Add(file);
            }

            if (existingFiles.Count == 0)
                return true;

            StringBuilder messsage = new StringBuilder();
            messsage.Append("OK to overwrite the following file(s)?");

            foreach (string existingFile in existingFiles)
            {
                messsage.Append("\n");
                messsage.Append(existingFile);
            }

            MessageBoxResult result = MessageBox.Show(messsage.ToString(), "Ok to overwrite?", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            return result == MessageBoxResult.Yes;
        }

        public static byte[] GetFileBytes(string path)
        {
            try
            {
                return File.ReadAllBytes(path);

            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(String.Format("Failed to read solution file from disk: {0}{1}{2}{3}{4}",
                    path, Environment.NewLine, ex.Message, Environment.NewLine, ex.StackTrace), MessageType.Error);

                return null;
            }
        }
    }
}