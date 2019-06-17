using D365DeveloperExtensions.Core;
using NLog;
using SolutionPackager.Resources;
using System;
using System.IO;
using System.Text;

namespace SolutionPackager
{
    public static class FileHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static string FormatSolutionVersionString(string solutionName, Version version, bool managed)
        {
            var result = new StringBuilder();

            result.Append($"{solutionName}_");

            result.Append(version.ToString().Replace(".", "_"));

            if (managed)
                result.Append("_managed");

            result.Append(".zip");

            return result.ToString();
        }

        public static string WriteTempFile(string filename, byte[] solutionBytes)
        {
            try
            {
                var tempFolder = Path.GetTempPath();

                var tempFile = Path.Combine(tempFolder, filename);
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
                File.WriteAllBytes(tempFile, solutionBytes);

                return tempFile;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorWritingTemp, ex);

                return null;
            }
        }

        public static DirectoryInfo CreateExtractFolder(string unmanagedPath)
        {
            try
            {
                var tempDirectory = Path.GetDirectoryName(unmanagedPath);
                if (Directory.Exists($"{tempDirectory}\\{Path.GetFileNameWithoutExtension(unmanagedPath)}"))
                    FileSystem.DeleteDirectory($"{tempDirectory}\\{Path.GetFileNameWithoutExtension(unmanagedPath)}");
                var extractedFolder =
                    Directory.CreateDirectory($"{tempDirectory}\\{Path.GetFileNameWithoutExtension(unmanagedPath)}");

                return extractedFolder;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorCreatingTemp, ex);

                return null;
            }
        }
    }
}