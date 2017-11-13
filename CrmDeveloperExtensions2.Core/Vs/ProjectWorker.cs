using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using VSLangProj;
using Constants = EnvDTE.Constants;

namespace CrmDeveloperExtensions2.Core.Vs
{
    public static class ProjectWorker
    {
        private static readonly string[] Extensions = { "HTM", "HTML", "CSS", "JS", "XML", "PNG", "JPG", "GIF", "XAP", "XSL", "XSLT", "ICO", "TS", "SVG", "RESX" };
        private static readonly string[] FolderExtensions = { "BUNDLE", "TT" };

        public static void ExcludeFolder(Project project, string folderName)
        {
            for (int i = 1; i <= project.ProjectItems.Count; i++)
            {
                Guid itemType = new Guid(project.ProjectItems.Item(i).Kind);
                if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                    continue;

                if (String.Equals(project.ProjectItems.Item(i).Name, folderName, StringComparison.CurrentCultureIgnoreCase))
                    project.ProjectItems.Item(i).Remove();
            }
        }

        public static string GetFolderProjectFileName(string projectFullName)
        {
            string path = Path.GetDirectoryName(projectFullName);
            if (path == null)
                return null;

            var dirName = new DirectoryInfo(path).Name;
            var fileName = new FileInfo(projectFullName).Name;
            string folderProjectFileName = dirName + "\\" + fileName;

            return folderProjectFileName;
        }

        public static bool IsProjectLoaded(Project project)
        {
            return string.Compare(Constants.vsProjectKindUnmodeled, project.Kind,
                       StringComparison.OrdinalIgnoreCase) != 0;
        }

        public static bool LoadProject(Project project)
        {
            object service = Package.GetGlobalService(typeof(IVsSolution));
            IVsSolution sln = (IVsSolution)service;
            IVsSolution4 sln4 = (IVsSolution4)service;

            int err = sln.GetProjectOfUniqueName(project.UniqueName, out IVsHierarchy hierarchy);
            if (VSConstants.S_OK != err)
                return false;

            const uint itemId = (uint)VSConstants.VSITEMID.Root;
            err = hierarchy.GetGuidProperty(itemId, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out Guid projectGuid);
            if (VSConstants.S_OK != err)
                return false;

            err = sln4.EnsureProjectIsLoaded(projectGuid, (uint)__VSBSLFLAGS.VSBSLFLAGS_None);
            return VSConstants.S_OK == err;
        }

        public static Project GetProjectByName(string projectName)
        {
            IList<Project> projects = GetProjects(false);
            foreach (Project project in projects)
            {
                if (project.Name != projectName) continue;

                return project;
            }

            return null;
        }

        public static IList<Project> GetProjects(bool excludeUnitTestProjects)
        {
            List<Project> list = new List<Project>();

            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return list;

            Projects projects = dte.Solution.Projects;

            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                if (project == null) continue;

                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(project));
                else
                    list.Add(project);
            }

            return excludeUnitTestProjects ? FilterUnitTestProjects(list) : list;
        }

        private static IList<Project> FilterUnitTestProjects(List<Project> list)
        {
            List<Project> filteredList = new List<Project>();
            foreach (Project project in list)
            {
                if (IsUnitTestProject(project))
                    continue;
                filteredList.Add(project);
            }

            return filteredList;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null) continue;

                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(subProject));
                else
                    list.Add(subProject);
            }
            return list;
        }

        public static ObservableCollection<string> GetProjectFolders(Project project)
        {
            ObservableCollection<string> folders = new ObservableCollection<string>();
            if (project == null)
                return folders;

            List<string> projectFolders = GetRootLevelProjectFolders(project);

            foreach (string projectFolder in projectFolders)
            {
                folders.Add(projectFolder);
            }

            return folders;
        }

        private static List<string> GetRootLevelProjectFolders(Project project)
        {
            List<string> projectFolders = new List<string>();
            var projectItems = project.ProjectItems;
            for (int i = 1; i <= projectItems.Count; i++)
            {
                var folders = GetFolders(projectItems.Item(i), String.Empty);
                foreach (string folder in folders)
                {
                    if (folder.ToUpper() == "/PROPERTIES") continue; //Don't add the project Properties folder
                    if (folder.ToUpper().StartsWith("/MY PROJECT")) continue; //Don't add the VB project My Project folders
                    projectFolders.Add(folder);
                }
            }

            projectFolders.Insert(0, "/");
            return projectFolders;
        }

        private static ObservableCollection<string> GetFolders(ProjectItem projectItem, string path)
        {
            ObservableCollection<string> projectFolders = new ObservableCollection<string>();
            if (new Guid(projectItem.Kind) != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            projectFolders.Add(path + "/" + projectItem.Name);
            for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
            {
                var folders = GetFolders(projectItem.ProjectItems.Item(i), path + "/" + projectItem.Name);
                foreach (string folder in folders)
                    projectFolders.Add(folder);
            }
            return projectFolders;
        }

        public static ObservableCollection<ComboBoxItem> GetProjectFilesForComboBox(Project project)
        {
            ObservableCollection<ComboBoxItem> projectFiles = new ObservableCollection<ComboBoxItem>();
            if (project == null)
                return projectFiles;

            var projectItems = project.ProjectItems;
            for (int i = 1; i <= projectItems.Count; i++)
            {
                var files = GetFiles(projectItems.Item(i), String.Empty);
                foreach (var comboBoxItem in files)
                    projectFiles.Add(comboBoxItem);
            }

            if (projectFiles.Count > 0)
                projectFiles.Insert(0, new ComboBoxItem { Content = String.Empty });

            return projectFiles;
        }

        private static ObservableCollection<ComboBoxItem> GetFiles(ProjectItem projectItem, string path)
        {
            ObservableCollection<ComboBoxItem> projectFiles = new ObservableCollection<ComboBoxItem>();
            if (projectItem.Kind.ToUpper() != "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") // VS Folder 
            {
                string ex = Path.GetExtension(projectItem.Name);
                if (ex == null || !Extensions.Contains(ex.Replace(".", String.Empty).ToUpper()) && !string.IsNullOrEmpty(ex) && !FolderExtensions.Contains(ex.Replace(".", String.Empty).ToUpper()))
                    return projectFiles;

                //Don't add file extensions that act as folders
                if (!FolderExtensions.Contains(ex.Replace(".", String.Empty).ToUpper()))
                    projectFiles.Add(new ComboBoxItem { Content = path + "/" + projectItem.Name, Tag = projectItem });

                if (projectItem.ProjectItems.Count <= 0)
                    return projectFiles;

                //Handle minified/bundled files that appear under other files in the project
                for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    ObservableCollection<ComboBoxItem> subFiles = GetFiles(projectItem.ProjectItems.Item(i), path);
                    foreach (ComboBoxItem comboBoxItem in subFiles)
                        projectFiles.Add(comboBoxItem);
                }
            }
            else
            {
                for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    //Ignore TypeScript typings folders
                    if (projectItem.Name.ToUpper() == "TYPINGS") continue;

                    var files = GetFiles(projectItem.ProjectItems.Item(i), path + "/" + projectItem.Name);
                    foreach (var comboBoxItem in files)
                        projectFiles.Add(comboBoxItem);
                }
            }
            return projectFiles;
        }

        public static string GetSdkCoreVersion(Project project)
        {
            if (!(project?.Object is VSProject vsproject))
                return null;

            foreach (Reference reference in vsproject.References)
            {
                if (reference.SourceProject != null)
                    continue;

                if (reference.Name == ExtensionConstants.MicrosoftXrmSdk)
                    return reference.Version;
            }

            return null;
        }

        public static string GetProjectTypeGuids(Project project)
        {
            string projectTypeGuids = String.Empty;

            object service = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            var solution = (IVsSolution)service;

            var result = solution.GetProjectOfUniqueName(project.UniqueName, out var hierarchy);

            if (result != 0)
                return projectTypeGuids;

            var aggregatableProject = (IVsAggregatableProject)hierarchy;
            aggregatableProject.GetAggregateProjectTypeGuids(out projectTypeGuids);

            return projectTypeGuids;
        }

        public static bool IsUnitTestProject(Project project)
        {
            string projectTypeGuids = GetProjectTypeGuids(project);
            if (string.IsNullOrEmpty(projectTypeGuids))
                return false;

            projectTypeGuids = projectTypeGuids.Replace("{", String.Empty).Replace("}", String.Empty);
            string[] guids = projectTypeGuids.Split(';');

            return guids.Contains(ExtensionConstants.UnitTestProjectType.ToString(), StringComparer.InvariantCultureIgnoreCase);
        }

        public static string GetAssemblyPath(Project project)
        {
            string fullPath = project.Properties.Item("FullPath").Value.ToString();
            string outputPath = project.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value.ToString();
            string outputDir = Path.Combine(fullPath, outputPath);
            string outputFileName = project.Properties.Item("OutputFileName").Value.ToString();
            string assemblyPath = Path.Combine(outputDir, outputFileName);

            return assemblyPath;
        }

        public static bool BuildProject(Project project)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return false;

            SolutionBuild solutionBuild = dte.Solution.SolutionBuild;
            solutionBuild.BuildProject(dte.Solution.SolutionBuild.ActiveConfiguration.Name, project.UniqueName, true);

            //0 = no errors
            return solutionBuild.LastBuildInfo <= 0;
        }

        public static string GetProjectPath(Project project)
        {
            string path = project.FullName;

            path = Path.GetDirectoryName(path);

            if (!string.IsNullOrEmpty(path))
                return path;

            OutputLogger.WriteToOutputWindow("Unable to get path from project", MessageType.Error);
            return null;
        }

        public static string GetOutputFile(Project project)
        {
            string outputFileName = project.Properties.Item("OutputFileName").Value.ToString();
            string path = GetOutputPath(project);
            return path == null ?
                null :
                Path.Combine(path, outputFileName);
        }

        public static string GetOutputPath(Project project)
        {
            ConfigurationManager configurationManager = project.ConfigurationManager;
            if (configurationManager == null) return null;

            Configuration activeConfiguration = configurationManager.ActiveConfiguration;
            string outputPath = activeConfiguration.Properties.Item("OutputPath").Value.ToString();
            string absoluteOutputPath = String.Empty;
            string projectFolder;

            if (outputPath.StartsWith(Path.DirectorySeparatorChar.ToString() + Path.DirectorySeparatorChar))
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.Length >= 2 && outputPath[0] == Path.VolumeSeparatorChar)
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.IndexOf("..\\", StringComparison.Ordinal) != -1)
            {
                projectFolder = Path.GetDirectoryName(project.FullName);

                while (outputPath.StartsWith("..\\"))
                {
                    outputPath = outputPath.Substring(3);
                    projectFolder = Path.GetDirectoryName(projectFolder);
                }

                if (projectFolder != null) absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }
            else
            {
                projectFolder = Path.GetDirectoryName(project.FullName);
                if (projectFolder != null)
                    absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }

            return absoluteOutputPath;
        }

        public static bool IsWorkflowProject(Project project)
        {
            if (!(project?.Object is VSProject vsproject))
                return false;

            foreach (Reference reference in vsproject.References)
            {
                if (reference.SourceProject != null)
                    continue;

                if (reference.Name == ExtensionConstants.MicrosoftXrmSdkWorkflow)
                    return true;
            }

            return false;
        }

        public static void AddProjectReference(VSProject vsproject, string referenceName)
        {
            try
            {
                Reference existingReference = vsproject.References.Find(referenceName);
                if (existingReference != null)
                    return;

                vsproject.References.Add(referenceName);
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    $"Failed to add refernce {referenceName}: {Environment.NewLine}{ex.Message}{Environment.NewLine}{ex.StackTrace}", MessageType.Error);
            }
        }
    }
}