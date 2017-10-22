using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using System;
using System.IO;
using System.Text;

namespace SolutionPackager
{
    public static class FileHandler
    {
        public static string FormatSolutionVersionString(string solutionName, Version version, bool managed)
        {
            StringBuilder result = new StringBuilder();

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
                OutputLogger.WriteToOutputWindow("Error Writing Solution To Temp Directory: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }

        public static DirectoryInfo CreateExtractFolder(string unmanagedPath)
        {
            try
            {
                string tempDirectory = Path.GetDirectoryName(unmanagedPath);
                if (Directory.Exists(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(unmanagedPath)))
                    Directory.Delete(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(unmanagedPath), true);
                DirectoryInfo extractedFolder =
                    Directory.CreateDirectory(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(unmanagedPath));

                return extractedFolder;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Creating Temp Directory To Extract Files: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                throw;
            }
        }
    }
}