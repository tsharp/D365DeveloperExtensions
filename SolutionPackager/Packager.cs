using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Microsoft.VisualStudio;
using System;
using System.IO;
using System.Text;

namespace SolutionPackager
{
    public static class Packager
    {
        public static bool CreatePackage(DTE dte, string toolPath, string solutionName, Project project,
            CrmDevExSolutionPackage crmDevExSolutionPackage, string fullPath, string commandArgs)
        {
            dte.ExecuteCommand($"shell {toolPath}", commandArgs);

            //Need this. Extend to allow bigger solutions to unpack
            //TODO: Better way?
            System.Threading.Thread.Sleep(5000);

            if (!crmDevExSolutionPackage.SaveSolutions)
                return true;

            project.ProjectItems.AddFromFile(fullPath);

            if (crmDevExSolutionPackage.CreateManaged)
                project.ProjectItems.AddFromFile(fullPath.Replace(".zip", "_managed.zip"));

            return true;
        }

        public static string GetPackageCommandArgs(Project project, string filename, string solutionProjectFolder, string fullPath, CrmDevExSolutionPackage crmDevExSolutionPackage)
        {
            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            if (string.IsNullOrEmpty(projectPath))
                return null;

            if (!CrmDeveloperExtensions2.Core.FileSystem.ConfirmOverwrite(
                new[] { fullPath, fullPath.Replace(".zip", "_managed.zip") }, true))
                return null;

            string commandArgs = CreatePackCommandArgs(projectPath, solutionProjectFolder, filename,
                crmDevExSolutionPackage.EnableSolutionPackagerLog, crmDevExSolutionPackage.CreateManaged);

            return commandArgs;
        }

        public static string GetExtractCommandArgs(string unmanagedZipPath, string managedZipPath, Project project, DirectoryInfo extractedFolder , 
            CrmDevExSolutionPackage crmDevExSolutionPackage)
        {
            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            if (string.IsNullOrEmpty(projectPath))
                return null;

            string commandArgs = CreateExtractCommandArgs(unmanagedZipPath, extractedFolder, projectPath,
                crmDevExSolutionPackage.EnableSolutionPackagerLog, crmDevExSolutionPackage.DownloadManaged);

            return commandArgs;
        }

        public static bool ExtractPackage(DTE dte, string toolPath, string unmanagedZipPath, string managedZipPath, Project project,
            CrmDevExSolutionPackage crmDevExSolutionPackage, DirectoryInfo extractedFolder, string commandArgs)
        {
            string solutionProjectFolder = GetProjectSolutionFolder(project, crmDevExSolutionPackage.ProjectFolder);

            dte.ExecuteCommand($"shell {toolPath}", commandArgs);

            //Need this. Extend to allow bigger solutions to unpack
            //TODO: Better way?
            System.Threading.Thread.Sleep(10000);

            bool solutionFileDelete = RemoveDeletedItems(extractedFolder.FullName, project.ProjectItems, crmDevExSolutionPackage.ProjectFolder);
            bool solutionFileAddChange = ProcessDownloadedSolution(extractedFolder, Path.GetDirectoryName(project.FullName), project.ProjectItems);

            Directory.Delete(extractedFolder.FullName, true);

            if (!crmDevExSolutionPackage.SaveSolutions)
                return true;

            //Solution change or file not present
            bool solutionChange = solutionFileDelete || solutionFileAddChange;
            bool solutionStored = StoreSolutionFile(unmanagedZipPath, project, solutionProjectFolder, solutionChange);
            if (crmDevExSolutionPackage.DownloadManaged && !string.IsNullOrEmpty(managedZipPath))
                solutionStored = StoreSolutionFile(managedZipPath, project, solutionProjectFolder, solutionChange);

            return solutionStored;
        }

        private static bool StoreSolutionFile(string zipPath, Project project, string solutionProjectFolder, bool solutionChange)
        {
            try
            {
                string filename = Path.GetFileName(zipPath);
                if (string.IsNullOrEmpty(filename))
                {
                    OutputLogger.WriteToOutputWindow("Error getting file name from temp path: " + zipPath, MessageType.Error);
                    return false;
                }

                string newSolutionFile = Path.Combine(solutionProjectFolder, filename);

                if (!solutionChange && File.Exists(newSolutionFile))
                    return true;

                if (File.Exists(newSolutionFile))
                    File.Delete(newSolutionFile);

                File.Move(zipPath, newSolutionFile);

                project.ProjectItems.AddFromFile(newSolutionFile);

                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error adding solution file to project: " + zipPath +
                    Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
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
                    if (FileEquals(Path.Combine(baseFolder, file.Name), file.FullName))
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

                if (CrmDeveloperExtensions2.Core.StringFormatting.FormatProjectKind(projectItem.Kind) == VSConstants.GUID_ItemType_PhysicalFile.ToString())
                {
                    name = Path.GetFileName(name);
                    // Do not delete the mapping file
                    if (name == CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerMapFile)
                        continue;
                    // Do not delete the config file
                    if (name == CrmDeveloperExtensions2.Core.Resources.Resource.ConfigFileName)
                        continue;
                    if (File.Exists(Path.Combine(extractedFolder, name)))
                        continue;

                    projectItem.Delete();
                    itemChanged = true;
                }

                if (CrmDeveloperExtensions2.Core.StringFormatting.FormatProjectKind(projectItem.Kind) == VSConstants.GUID_ItemType_PhysicalFolder.ToString())
                {
                    name = new DirectoryInfo(name).Name;
                    if (name == projectFolder || name == "Properties")
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

        public static string GetProjectSolutionFolder(Project project, string projectFolder)
        {
            if (projectFolder.StartsWith("/"))
                projectFolder = projectFolder.Substring(1);

            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            projectPath = Path.Combine(projectPath, projectFolder);

            if (!Directory.Exists(projectPath))
                project.ProjectItems.AddFolder(projectFolder);

            return projectPath;
        }

        private static bool FileEquals(string path1, string path2)
        {
            FileInfo first = new FileInfo(path1);
            FileInfo second = new FileInfo(path2);

            if (first.Length != second.Length)
                return false;

            int iterations = (int)Math.Ceiling((double)first.Length / sizeof(Int64));

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                byte[] one = new byte[sizeof(Int64)];
                byte[] two = new byte[sizeof(Int64)];

                for (int i = 0; i < iterations; i++)
                {
                    fs1.Read(one, 0, sizeof(Int64));
                    fs2.Read(two, 0, sizeof(Int64));

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                        return false;
                }
            }

            return true;
        }

        public static string CreateToolPath(DTE dte)
        {
            CrmDeveloperExtensions2.Core.UserOptionsGrid.GetSolutionPackagerToolPath(dte);
            string spPath = CrmDeveloperExtensions2.Core.UserOptionsGrid.GetSolutionPackagerToolPath(dte);

            if (string.IsNullOrEmpty(spPath))
            {
                OutputLogger.WriteToOutputWindow("Please set the Solution Packager path in options", MessageType.Error);
                return null;
            }

            if (!spPath.EndsWith("\\"))
                spPath += "\\";

            string toolPath = @"""" + spPath + "SolutionPackager.exe" + @"""";

            if (!File.Exists(spPath + "SolutionPackager.exe"))
            {
                OutputLogger.WriteToOutputWindow($"SolutionPackager.exe not found at: {spPath}", MessageType.Error);
                return null;
            }

            return toolPath;
        }

        private static string CreatePackCommandArgs(string projectPath, string solutionProjectFolder, string filename,
            bool enableSolutionPackagerLogging, bool downloadManaged)
        {
            StringBuilder command = new StringBuilder();
            command.Append(" /action: Pack");
            command.Append($" /zipfile: \"{Path.Combine(solutionProjectFolder, filename)}\"");
            command.Append($" /folder: \"{projectPath}\"");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (File.Exists(Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerMapFile)))
                command.Append($" /map: \"{Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerMapFile)}\"");

            // Write Solution Package output to a log file named SolutionPackager.log in the root folder of the project
            if (enableSolutionPackagerLogging)
                command.Append($" /log: \"{Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerLogFile)}\"");

            // Pack managed solution as well.
            if (downloadManaged)
                command.Append(" /packagetype:Both");

            return command.ToString();
        }

        private static string CreateExtractCommandArgs(string unmanagedZipPath, DirectoryInfo extractFolder, string projectPath, bool enableSolutionPackagerLogging, bool downloadManaged)
        {
            StringBuilder command = new StringBuilder();
            command.Append(" /action: Extract");
            command.Append($" /zipfile: \"{unmanagedZipPath}\"");
            command.Append($" /folder: \"{extractFolder.FullName}\"");
            command.Append(" /clobber");

            // Add a mapping file which should be in the root folder of the project and be named mapping.xml
            if (File.Exists(Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerMapFile)))
                command.Append($" /map: \"{Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerMapFile)}\"");

            // Write Solution Package output to a log file named SolutionPackager.log in the root folder of the project
            if (enableSolutionPackagerLogging)
                command.Append($" /log: \"{Path.Combine(projectPath, CrmDeveloperExtensions2.Core.ExtensionConstants.SolutionPackagerLogFile)}\"");

            // Unpack managed solution as well.
            if (downloadManaged)
                command.Append(" /packagetype:Both");

            return command.ToString();
        }
    }
}