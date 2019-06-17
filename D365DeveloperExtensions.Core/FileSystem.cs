using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace D365DeveloperExtensions.Core
{
    public static class FileSystem
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool IsDirectoryEmpty(string path)
        {
            var items = Directory.EnumerateFileSystemEntries(path);
            using (var en = items.GetEnumerator())
                return !en.MoveNext();
        }

        public static DirectoryInfo GetDirectory(string input)
        {
            var path = Path.GetDirectoryName(input);
            if (path == null)
                throw new Exception(Resource.ErrorMessage_DirectoryFromString);

            var directory = new DirectoryInfo(path);
            if (!directory.Exists)
                throw new Exception(Resource.ErrorMessage_DirectoryFromString);

            return directory;
        }

        public static string WriteTempFile(string name, byte[] content)
        {
            try
            {
                var tempFolder = Path.GetTempPath();
                var fileName = Path.GetFileName(name);
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
                MessageBox.Show(Resource.ErrorMessage_WriteTempFile);
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
                MessageBox.Show(Resource.ErrorMessage_WriteFile);
                throw;
            }
        }

        public static string BoundFileToLocalPath(string boundFile, string projectPath)
        {
            if (string.IsNullOrEmpty(boundFile))
                return "";

            var path = File.Exists(projectPath)
                ? Path.GetDirectoryName(projectPath)
                : projectPath;

            if (path == null)
                return null;

            if (boundFile.StartsWith("/"))
                boundFile = boundFile.Substring(1);

            return Path.Combine(path, boundFile.Replace("/", "\\"));
        }

        public static string LocalPathToCrmPath(string projectPath, string filename)
        {
            var newName = filename.Replace(projectPath, string.Empty).Replace("\\", "/");
            if (!newName.StartsWith("/"))
                newName = "/" + newName;
            return newName;
        }

        public static bool DoesFileExist(string[] files, bool checkAll)
        {
            foreach (var file in files)
            {
                var exists = File.Exists(file);
                if (exists && !checkAll)
                    return true;

                if (!exists && checkAll)
                    return false;
            }

            return checkAll;
        }

        public static void RenameFile(string path)
        {
            try
            {
                File.Move(path, $"{path}.{DateTime.Now:MMddyyyyHHmmss}");
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_UnableRenameFile}: {path}", ex);
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_UnableRenameFile, MessageType.Error);
                throw;
            }
        }

        public static bool ConfirmOverwrite(string[] files, bool checkAll)
        {
            var existingFiles = new List<string>();

            foreach (var file in files)
            {
                var exists = DoesFileExist(new[] { file }, true);
                if (exists)
                    existingFiles.Add(file);
            }

            if (existingFiles.Count == 0)
                return true;

            var message = new StringBuilder();
            message.Append(Resource.ConfirmMessage_OverwriteFiles);

            foreach (var existingFile in existingFiles)
            {
                message.Append("\n");
                message.Append(existingFile);
            }

            var result = MessageBox.Show(message.ToString(), Resource.ConfirmMessage_OverwriteFiles_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

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
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_ReadFile}: {path}", ex);
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_ReadFile, MessageType.Error);
                return null;
            }
        }

        public static string GetFileText(string path)
        {
            try
            {
                return File.ReadAllText(path);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_ReadFile}: {path}", ex);
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_ReadFile, MessageType.Error);
                return null;
            }
        }

        public static bool FileEquals(string path1, string path2)
        {
            var first = new FileInfo(path1);
            var second = new FileInfo(path2);

            if (first.Length != second.Length)
                return false;

            var iterations = (int)Math.Ceiling((double)first.Length / sizeof(Int64));

            using (var fs1 = first.OpenRead())
            using (var fs2 = second.OpenRead())
            {
                var one = new byte[sizeof(Int64)];
                var two = new byte[sizeof(Int64)];

                for (var i = 0; i < iterations; i++)
                {
                    fs1.Read(one, 0, sizeof(Int64));
                    fs2.Read(two, 0, sizeof(Int64));

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                        return false;
                }
            }

            return true;
        }

        public static bool IsValidFilename(string fileName, string sourceFolder)
        {
            return !string.IsNullOrEmpty(fileName) &&
                          fileName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 &&
                          !File.Exists(Path.Combine(sourceFolder, fileName));
        }

        public static bool IsValidFolder(string folder)
        {
            if (string.IsNullOrEmpty(folder))
                return true;

            try
            {
                var di = new DirectoryInfo(folder);
                return di.Exists;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void DeleteDirectory(string path)
        {
            foreach (var directory in Directory.GetDirectories(path))
            {
                DeleteDirectory(directory);
            }

            try
            {
                Thread.Sleep(100);
                Directory.Delete(path, true);
            }
            catch (IOException)
            {
                Directory.Delete(path, true);
            }
            catch (UnauthorizedAccessException)
            {
                Directory.Delete(path, true);
            }
        }
    }
}