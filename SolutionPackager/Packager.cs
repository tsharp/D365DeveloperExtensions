using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.UserOptions;
using EnvDTE;
using Microsoft.VisualStudio;
using NLog;
using SolutionPackager.Models;
using SolutionPackager.Resources;
using System;
using System.IO;
using System.Text;

namespace SolutionPackager
{
    public static class Packager
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool CreatePackage(DTE dte, string toolPath, PackSettings packSettings, string commandArgs)
        {
            dte.ExecuteCommand($"shell {toolPath}", commandArgs);

            //Need this. Extend to allow bigger solutions to unpack
            //TODO: Find a better way
            System.Threading.Thread.Sleep(5000);

            if (!packSettings.SaveSolutions)
                return true;

            packSettings.Project.ProjectItems.AddFromFile(packSettings.FullFilePath);

            return true;
        }

        public static string GetPackageCommandArgs(PackSettings packSettings)
        {
            if (!FileSystem.ConfirmOverwrite(
                new[] { packSettings.FullFilePath, packSettings.FullFilePath.Replace(".zip", "_managed.zip") }, true))
                return null;

            string commandArgs = CreatePackCommandArgs(packSettings);

            return commandArgs;
        }

        public static string GetExtractCommandArgs(UnpackSettings unpackSettings)
        {
            string commandArgs = CreateExtractCommandArgs(unpackSettings);

            return commandArgs;
        }

        public static bool ExtractPackage(DTE dte, string toolPath, UnpackSettings unpackSettings, string commandArgs)
        {
            dte.ExecuteCommand($"shell {toolPath}", commandArgs);

            //Need this. Extend to allow bigger solutions to unpack
            //TODO: Find a better way
            System.Threading.Thread.Sleep(10000);

            bool solutionFileDelete = RemoveDeletedItems(unpackSettings.ExtractedFolder.FullName, unpackSettings.Project.ProjectItems,
                unpackSettings.SolutionPackageConfig.packagepath);
            bool solutionFileAddChange = ProcessDownloadedSolution(unpackSettings.ExtractedFolder, unpackSettings.ProjectPackageFolder,
                unpackSettings.Project.ProjectItems);

            Directory.Delete(unpackSettings.ExtractedFolder.FullName, true);

            if (!unpackSettings.SaveSolutions)
                return true;

            //Solution change or file not present
            bool solutionChange = solutionFileDelete || solutionFileAddChange;
            bool solutionStored = StoreSolutionFile(unpackSettings, solutionChange);

            return solutionStored;
        }

        private static bool StoreSolutionFile(UnpackSettings unpackSettings, bool solutionChange)
        {
            try
            {
                string filename = Path.GetFileName(unpackSettings.DownloadedZipPath);
                if (string.IsNullOrEmpty(filename))
                {
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_ErrorGettingFileNameFromTemp}: {unpackSettings.DownloadedZipPath}", MessageType.Error);
                    return false;
                }

                string newSolutionFile = Path.Combine(unpackSettings.ProjectSolutionFolder, filename);

                if (!solutionChange && File.Exists(newSolutionFile))
                    return true;

                if (File.Exists(newSolutionFile))
                    File.Delete(newSolutionFile);

                File.Move(unpackSettings.DownloadedZipPath, newSolutionFile);

                unpackSettings.Project.ProjectItems.AddFromFile(newSolutionFile);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorAddingSolutionFileProject, ex);

                return false;
            }
        }

        private static bool ProcessDownloadedSolution(DirectoryInfo extractedFolder, string baseFolder, ProjectItems projectItems)
        {
            bool itemChanged = false;

            //Handle file adds
            foreach (FileInfo file in extractedFolder.GetFiles())
            {
                if (File.Exists(Path.Combine(baseFolder, file.Name)))
                {
                    if (FileSystem.FileEquals(Path.Combine(baseFolder, file.Name), file.FullName))
                        continue;
                }

                File.Copy(file.FullName, Path.Combine(baseFolder, file.Name), true);
                projectItems.AddFromFile(Path.Combine(baseFolder, file.Name));
                itemChanged = true;
            }

            //Handle folder adds
            foreach (DirectoryInfo folder in extractedFolder.GetDirectories())
            {
                if (!Directory.Exists(Path.Combine(baseFolder, folder.Name)))
                    Directory.CreateDirectory(Path.Combine(baseFolder, folder.Name));

                var newProjectItems = projectItems;
                bool subItemChanged = ProcessDownloadedSolution(folder, Path.Combine(baseFolder, folder.Name), newProjectItems);
                if (subItemChanged)
                    itemChanged = true;
            }

            return itemChanged;
        }

        private static bool RemoveDeletedItems(string extractedFolder, ProjectItems projectItems, string projectFolder)
        {
            bool itemChanged = false;

            //Handle file & folder deletes
            foreach (ProjectItem projectItem in projectItems)
            {
                string name = projectItem.FileNames[0];
                if (string.IsNullOrEmpty(name))
                    continue;

                if (StringFormatting.RemoveBracesToUpper(projectItem.Kind) == VSConstants.GUID_ItemType_PhysicalFile.ToString())
                {
                    name = Path.GetFileName(name);
                    // Do not delete the mapping file
                    if (name == ExtensionConstants.SolutionPackagerMapFile)
                        continue;
                    // Do not delete the config file
                    if (name == ExtensionConstants.SpklConfigFile)
                        continue;
                    if (File.Exists(Path.Combine(extractedFolder, name)))
                        continue;

                    projectItem.Delete();
                    itemChanged = true;
                }

                if (StringFormatting.RemoveBracesToUpper(projectItem.Kind) == VSConstants.GUID_ItemType_PhysicalFolder.ToString())
                {
                    name = new DirectoryInfo(name).Name;
                    if (name == projectFolder || name.Equals(CrmDeveloperExtensions2.Core.Resources.Resource.Constant_PropertiesFolder,
                        StringComparison.CurrentCultureIgnoreCase))
                        continue;

                    if (!Directory.Exists(Path.Combine(extractedFolder, name)))
                    {
                        projectItem.Delete();
                        itemChanged = true;
                    }
                    else
                    {
                        if (projectItem.ProjectItems.Count <= 0)
                            continue;

                        bool subItemChanged = RemoveDeletedItems(Path.Combine(extractedFolder, name),
                            projectItem.ProjectItems, projectFolder);
                        if (subItemChanged)
                            itemChanged = true;
                    }
                }
            }

            return itemChanged;
        }

        public static string GetProjectPackageFolder(Project project, string projectFolder)
        {
            if (projectFolder.StartsWith("/", StringComparison.CurrentCultureIgnoreCase))
                projectFolder = projectFolder.Substring(1);

            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            projectPath = Path.Combine(projectPath, projectFolder);

            if (!Directory.Exists(projectPath))
                project.ProjectItems.AddFolder(projectFolder);

            return projectPath;
        }

        public static string CreateToolPath()
        {
            string spPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.SolutionPackagerToolPath);
            if (string.IsNullOrEmpty(spPath))
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_SetSolutionPackagerPath, MessageType.Error);
                return null;
            }

            if (!spPath.EndsWith("\\", StringComparison.CurrentCultureIgnoreCase))
                spPath += "\\";

            string toolPath = @"""" + spPath + "SolutionPackager.exe" + @"""";

            if (File.Exists(spPath + "SolutionPackager.exe"))
                return toolPath;

            OutputLogger.WriteToOutputWindow($"S{Resource.ErrorMessage_SolutionPackagerNotFound}: {spPath}", MessageType.Error);
            return null;
        }

        private static string CreatePackCommandArgs(PackSettings packSettings)
        {
            var zipFolder = packSettings.SaveSolutions
                ? packSettings.ProjectSolutionFolder
                : packSettings.ProjectPackageFolder;

            StringBuilder command = new StringBuilder();
            command.Append(" /action: Pack");
            command.Append($" /zipfile: \"{Path.Combine(zipFolder, packSettings.FileName)}\"");
            command.Append($" /folder: \"{packSettings.ProjectPackageFolder}\"");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (packSettings.SolutionPackageConfig.map != null)
            {
                MapFile.Create(packSettings.SolutionPackageConfig, Path.Combine(packSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile));
                if (File.Exists(Path.Combine(packSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile)))
                    command.Append($" /map: \"{Path.Combine(packSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile)}\"");
            }

            // Write Solution Package output to a log file named SolutionPackager.log in the root folder of the project
            if (packSettings.EnablePackagerLogging)
                command.Append($" /log: \"{Path.Combine(packSettings.ProjectPath, ExtensionConstants.SolutionPackagerLogFile)}\"");

            // Pack managed or unmanaged
            command.Append($" /packagetype:{packSettings.SolutionPackageConfig.packagetype}");

            return command.ToString();
        }

        private static string CreateExtractCommandArgs(UnpackSettings unpackSettings)
        {
            StringBuilder command = new StringBuilder();
            command.Append(" /action: Extract");
            command.Append($" /zipfile: \"{unpackSettings.DownloadedZipPath}\"");
            command.Append($" /folder: \"{unpackSettings.ExtractedFolder.FullName}\"");
            command.Append(" /clobber");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (unpackSettings.SolutionPackageConfig.map != null)
            {
                MapFile.Create(unpackSettings.SolutionPackageConfig, Path.Combine(unpackSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile));
                if (File.Exists(Path.Combine(unpackSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile)))
                    command.Append($" /map: \"{Path.Combine(unpackSettings.ProjectPath, ExtensionConstants.SolutionPackagerMapFile)}\"");
            }

            // Write Solution Package output to a log file named SolutionPackager.log in the root folder of the project
            if (unpackSettings.EnablePackagerLogging)
                command.Append($" /log: \"{Path.Combine(unpackSettings.ProjectPath, ExtensionConstants.SolutionPackagerLogFile)}\"");

            // Unpack managed or unmanaged
            command.Append($" /packagetype:{unpackSettings.SolutionPackageConfig.packagetype}");

            return command.ToString();
        }
    }
}