using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using Microsoft.VisualStudio;
using NLog;
using SolutionPackager.Models;
using SolutionPackager.Resources;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;

namespace SolutionPackager
{
    public static class Packager
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool CreatePackage(string toolPath, PackSettings packSettings, string commandArgs)
        {
            var command = new SolutionPackagerCommand
            {
                Action = SolutionPackagerAction.Pack.ToString(),
                CommandArgs = commandArgs,
                ToolPath = toolPath,
                SolutionName = packSettings.CrmSolution.Name
            };

            ExecuteSolutionPackager(command);

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

            var commandArgs = CreatePackCommandArgs(packSettings);

            return commandArgs;
        }

        public static string GetExtractCommandArgs(UnpackSettings unpackSettings)
        {
            var commandArgs = CreateExtractCommandArgs(unpackSettings);

            return commandArgs;
        }

        public static bool ExtractPackage(DTE dte, string toolPath, UnpackSettings unpackSettings, string commandArgs)
        {
            var command = new SolutionPackagerCommand
            {
                Action = SolutionPackagerAction.Extract.ToString(),
                CommandArgs = commandArgs,
                ToolPath = toolPath,
                SolutionName = unpackSettings.CrmSolution.Name
            };

            ExecuteSolutionPackager(command);

            var solutionFileDelete = RemoveDeletedItems(unpackSettings.ExtractedFolder.FullName,
                unpackSettings.Project.ProjectItems,
                unpackSettings.SolutionPackageConfig.packagepath);
            var solutionFileAddChange = ProcessDownloadedSolution(unpackSettings.ExtractedFolder,
                unpackSettings.ProjectPackageFolder,
                unpackSettings.Project.ProjectItems);

            FileSystem.DeleteDirectory(unpackSettings.ExtractedFolder.FullName);

            if (!unpackSettings.SaveSolutions)
                return true;

            //Solution change or file not present
            var solutionChange = solutionFileDelete || solutionFileAddChange;
            var solutionStored = StoreSolutionFile(unpackSettings, solutionChange);

            return solutionStored;
        }

        private static bool StoreSolutionFile(UnpackSettings unpackSettings, bool solutionChange)
        {
            try
            {
                var filename = Path.GetFileName(unpackSettings.DownloadedZipPath);
                if (string.IsNullOrEmpty(filename))
                {
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_ErrorGettingFileNameFromTemp}: {unpackSettings.DownloadedZipPath}", MessageType.Error);
                    return false;
                }

                var newSolutionFile = Path.Combine(unpackSettings.ProjectSolutionFolder, filename);

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
            var itemChanged = false;

            //Handle file adds
            foreach (var file in extractedFolder.GetFiles())
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
            foreach (var folder in extractedFolder.GetDirectories())
            {
                if (!Directory.Exists(Path.Combine(baseFolder, folder.Name)))
                    Directory.CreateDirectory(Path.Combine(baseFolder, folder.Name));

                var newProjectItems = projectItems;
                var subItemChanged = ProcessDownloadedSolution(folder, Path.Combine(baseFolder, folder.Name), newProjectItems);
                if (subItemChanged)
                    itemChanged = true;
            }

            return itemChanged;
        }

        private static bool RemoveDeletedItems(string extractedFolder, ProjectItems projectItems, string projectFolder)
        {
            var itemChanged = false;

            //Handle file & folder deletes
            foreach (ProjectItem projectItem in projectItems)
            {
                var name = projectItem.FileNames[0];
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
                    if (name == projectFolder || name.Equals(D365DeveloperExtensions.Core.Resources.Resource.Constant_PropertiesFolder,
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

                        var subItemChanged = RemoveDeletedItems(Path.Combine(extractedFolder, name),
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

            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            projectPath = Path.Combine(projectPath, projectFolder);

            if (!Directory.Exists(projectPath))
                project.ProjectItems.AddFolder(projectFolder);

            return projectPath;
        }

        public static string CreateToolPath()
        {
            var spPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.SolutionPackagerToolPath);
            if (string.IsNullOrEmpty(spPath))
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_SetSolutionPackagerPath, MessageType.Error);
                return null;
            }

            if (!spPath.EndsWith("\\", StringComparison.CurrentCultureIgnoreCase))
                spPath += "\\";

            var toolPath = @"""" + spPath + "SolutionPackager.exe" + @"""";

            if (File.Exists(spPath + "SolutionPackager.exe"))
                return toolPath;

            OutputLogger.WriteToOutputWindow($"S{Resource.ErrorMessage_SolutionPackagerNotFound}: {spPath}", MessageType.Error);
            return null;
        }

        private static string CreatePackCommandArgs(PackSettings packSettings)
        {
            var command = new StringBuilder();
            command.Append(" /action: Pack");
            command.Append($" /zipfile: \"{Path.Combine(packSettings.ProjectSolutionFolder, packSettings.FileName)}\"");
            command.Append($" /folder: \"{packSettings.ProjectPackageFolder}\"");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (packSettings.SolutionPackageConfig.map != null && packSettings.UseMapFile)
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
            var command = new StringBuilder();
            command.Append(" /action: Extract");
            command.Append($" /zipfile: \"{unpackSettings.DownloadedZipPath}\"");
            command.Append($" /folder: \"{unpackSettings.ExtractedFolder.FullName}\"");
            command.Append(" /clobber");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (unpackSettings.SolutionPackageConfig.map != null && unpackSettings.UseMapFile)
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

        public static bool ExecuteSolutionPackager(SolutionPackagerCommand command)
        {
            OutputLogger.WriteToOutputWindow($"{Resource.Message_Begin} {command.Action}: {command.SolutionName}", MessageType.Info);

            const int timeout = 60000;
            var workingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(workingDirectory))
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_CouldNotSetWorkingDirectory, MessageType.Error);
                return false;
            }

            using (var process = new System.Diagnostics.Process())
            {
                var processStartInfo = CreateProcessStartInfo(command);
                process.StartInfo = processStartInfo;
                process.StartInfo.WorkingDirectory = workingDirectory;

                var output = new StringBuilder();
                var errorDataReceived = new StringBuilder();

                using (var outputWaitHandle = new AutoResetEvent(false))
                {
                    using (var errorWaitHandle = new AutoResetEvent(false))
                    {
                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (e.Data == null)
                                outputWaitHandle.Set();
                            else
                                output.AppendLine(e.Data);
                        };
                        process.ErrorDataReceived += (sender, e) =>
                        {
                            if (e.Data == null)
                                errorWaitHandle.Set();
                            else
                                errorDataReceived.AppendLine(e.Data);
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();

                        string message;
                        if (process.WaitForExit(timeout) && outputWaitHandle.WaitOne(timeout) && errorWaitHandle.WaitOne(timeout))
                        {
                            if (process.ExitCode == 0)
                            {
                                OutputLogger.WriteToOutputWindow($"{Resource.Message_End} {command.Action}: {command.SolutionName}", MessageType.Info);
                                return true;
                            }

                            message = $"{Resource.Message_ErrorExecutingSolutionPackager}: {command.Action}: {command.SolutionName}";
                        }
                        else
                        {
                            message = $"{Resource.Message_TimoutExecutingSolutionPackager}: {command.Action}: {command.SolutionName}";
                        }

                        ExceptionHandler.LogProcessError(Logger, message, errorDataReceived.ToString());
                        MessageBox.Show(message);
                    }
                }
            }

            return false;
        }

        private static ProcessStartInfo CreateProcessStartInfo(SolutionPackagerCommand command)
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/c \"{command.ToolPath} {command.CommandArgs}\"",
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                UseShellExecute = false
            };

            return processStartInfo;
        }
    }
}