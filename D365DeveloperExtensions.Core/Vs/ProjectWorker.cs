using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using EnvDTE80;
using Microsoft.Build.Evaluation;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using VSLangProj;
using Constants = EnvDTE.Constants;
using Project = EnvDTE.Project;
using ProjectItem = EnvDTE.ProjectItem;

namespace D365DeveloperExtensions.Core.Vs
{
    public static class ProjectWorker
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly List<string> Extensions = new List<string> { "HTM", "HTML", "CSS", "JS", "XML", "PNG", "JPG", "GIF", "XAP", "XSL", "XSLT", "ICO", "SVG", "RESX", "MAP", "TS" };
        private static readonly string[] FolderExtensions = { "BUNDLE", "TT" };
        private static readonly string[] IgnoreFolders = { "/TYPINGS", "/NODE_MODULES" };
        private static readonly string[] IgnoreFiles = { "GULPFILE.JS", "KARMA.CONF.JS", "PROTRACTOR.CONF.JS", ".D.TS" };

        public static void ExcludeFolder(Project project, string folderName)
        {
            for (var i = 1; i <= project.ProjectItems.Count; i++)
            {
                var itemType = new Guid(project.ProjectItems.Item(i).Kind);
                if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                    continue;

                if (string.Equals(project.ProjectItems.Item(i).Name, folderName, StringComparison.CurrentCultureIgnoreCase))
                    project.ProjectItems.Item(i).Remove();
            }
        }

        public static string GetFolderProjectFileName(string projectFullName)
        {
            var path = Path.GetDirectoryName(projectFullName);
            if (path == null)
                return null;

            var dirName = new DirectoryInfo(path).Name;
            var fileName = new FileInfo(projectFullName).Name;
            var folderProjectFileName = $"{dirName}\\{fileName}";

            return folderProjectFileName;
        }

        public static bool IsProjectLoaded(Project project)
        {
            return string.Compare(Constants.vsProjectKindUnmodeled, project.Kind,
                       StringComparison.OrdinalIgnoreCase) != 0;
        }

        public static bool LoadProject(Project project)
        {
            var service = Package.GetGlobalService(typeof(IVsSolution));
            var sln = (IVsSolution)service;
            var sln4 = (IVsSolution4)service;

            var err = sln.GetProjectOfUniqueName(project.UniqueName, out var hierarchy);
            if (VSConstants.S_OK != err)
                return false;

            const uint itemId = (uint)VSConstants.VSITEMID.Root;
            err = hierarchy.GetGuidProperty(itemId, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out var projectGuid);
            if (VSConstants.S_OK != err)
                return false;

            err = sln4.EnsureProjectIsLoaded(projectGuid, (uint)__VSBSLFLAGS.VSBSLFLAGS_None);
            return VSConstants.S_OK == err;
        }

        public static Project GetProjectByName(string projectName)
        {
            var projects = GetProjects(false);
            foreach (var project in projects)
            {
                if (project.Name != projectName) continue;

                return project;
            }

            return null;
        }

        public static IList<Project> GetProjects(bool excludeUnitTestProjects)
        {
            var list = new List<Project>();

            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return list;

            var projects = dte.Solution.Projects;

            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                if (!(item.Current is Project project))
                    continue;

                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(project));
                else
                {
                    if (project.Kind != ExtensionConstants.VsSharedProject && project.Kind != Constants.vsProjectKindMisc)
                        list.Add(project);
                }
            }

            return excludeUnitTestProjects ? FilterUnitTestProjects(list) : list;
        }

        private static IList<Project> FilterUnitTestProjects(List<Project> list)
        {
            var filteredList = new List<Project>();
            foreach (var project in list)
            {
                if (IsUnitTestProject(project))
                    continue;

                filteredList.Add(project);
            }

            return filteredList;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            var list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                    continue;

                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    list.AddRange(GetSolutionFolderProjects(subProject));
                else
                    list.Add(subProject);
            }
            return list;
        }

        public static ObservableCollection<string> GetProjectFolders(Project project, ProjectType projectType)
        {
            var folders = new ObservableCollection<string>();
            if (project == null)
                return folders;

            var projectFolders = GetRootLevelProjectFolders(project);

            foreach (var projectFolder in projectFolders)
            {
                folders.Add(projectFolder);
            }

            if (projectType != ProjectType.WebResource)
                return folders;

            var removeFolders = new List<string>();
            foreach (var ignoreFolder in IgnoreFolders)
            {
                removeFolders.AddRange(folders.Where(f => f.StartsWith(ignoreFolder, StringComparison.InvariantCultureIgnoreCase)).ToList());
            }

            foreach (var removeFolder in removeFolders)
            {
                folders.Remove(removeFolder);
            }

            return folders;
        }

        private static List<string> GetRootLevelProjectFolders(Project project)
        {
            var projectFolders = new List<string>();
            var projectItems = project.ProjectItems;
            for (var i = 1; i <= projectItems.Count; i++)
            {
                var folders = GetFolders(projectItems.Item(i), string.Empty);
                foreach (var folder in folders)
                {
                    if (folder.ToUpper() == $"/{Resource.Constant_PropertiesFolder.ToUpper()}") continue; //Don't add the project Properties folder
                    if (folder.ToUpper().StartsWith($"/{Resource.Constant_MyProjectFolder.ToUpper()}")) continue; //Don't add the VB project My Project folders
                    projectFolders.Add(folder);
                }
            }

            projectFolders.Insert(0, "/");
            return projectFolders;
        }

        private static ObservableCollection<string> GetFolders(ProjectItem projectItem, string path)
        {
            var projectFolders = new ObservableCollection<string>();
            if (new Guid(projectItem.Kind) != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            projectFolders.Add($"{path}/{projectItem.Name}");
            for (var i = 1; i <= projectItem.ProjectItems.Count; i++)
            {
                var folders = GetFolders(projectItem.ProjectItems.Item(i), $"{path}/{projectItem.Name}");
                foreach (var folder in folders)
                    projectFolders.Add(folder);
            }
            return projectFolders;
        }

        public static ObservableCollection<ComboBoxItem> GetProjectFilesForComboBox(Project project, bool hasTsConfig)
        {
            SetProjectExtensions(hasTsConfig);

            var projectFiles = new ObservableCollection<ComboBoxItem>();
            if (project == null)
                return projectFiles;

            var projectItems = project.ProjectItems;
            for (var i = 1; i <= projectItems.Count; i++)
            {
                var files = GetFiles(projectItems.Item(i), string.Empty);
                foreach (var comboBoxItem in files)
                    projectFiles.Add(comboBoxItem);
            }

            if (projectFiles.Count > 0)
                projectFiles.Insert(0, new ComboBoxItem { Content = string.Empty });

            //Remove files that are in ignored folders
            var copy = new ObservableCollection<ComboBoxItem>(projectFiles);
            foreach (var comboBoxItem in copy)
            {
                foreach (var ignoreFolder in IgnoreFolders)
                {
                    if (comboBoxItem.Content.ToString().StartsWith(ignoreFolder, StringComparison.InvariantCultureIgnoreCase))
                        projectFiles.Remove(comboBoxItem);
                }
            }

            //Remove ignored files
            copy = new ObservableCollection<ComboBoxItem>(projectFiles);
            foreach (var comboBoxItem in copy)
            {
                foreach (var ignoreFile in IgnoreFiles)
                {
                    if (comboBoxItem.Content.ToString().ToUpper().EndsWith(ignoreFile))
                        projectFiles.Remove(comboBoxItem);
                }
            }

            return projectFiles;
        }

        private static void SetProjectExtensions(bool hasTsConfig)
        {
            if (!hasTsConfig)
            {
                Extensions.Remove("MAP");
                Extensions.Remove("TS");
            }
            else
            {
                if (!Extensions.Contains("MAP"))
                    Extensions.Add("MAP");
                if (!Extensions.Contains("TS"))
                    Extensions.Add("TS");
            }
        }

        private static ObservableCollection<ComboBoxItem> GetFiles(ProjectItem projectItem, string path)
        {
            var projectFiles = new ObservableCollection<ComboBoxItem>();
            if (new Guid(projectItem.Kind) != VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var ex = Path.GetExtension(projectItem.Name);
                if (ex == null || !Extensions.Contains(ex.Replace(".", string.Empty).ToUpper()) &&
                    !string.IsNullOrEmpty(ex) && !FolderExtensions.Contains(ex.Replace(".", string.Empty).ToUpper()))
                    return projectFiles;

                //Don't add file extensions that act as folders
                if (!FolderExtensions.Contains(ex.Replace(".", string.Empty).ToUpper()))
                    projectFiles.Add(new ComboBoxItem { Content = $"{path}/{projectItem.Name}", Tag = projectItem });

                if (projectItem.ProjectItems == null || projectItem.ProjectItems.Count <= 0)
                    return projectFiles;

                //Handle minified/bundled files that appear under other files in the project
                for (var i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    var subFiles = GetFiles(projectItem.ProjectItems.Item(i), path);
                    foreach (var comboBoxItem in subFiles)
                        projectFiles.Add(comboBoxItem);
                }
            }
            else
            {
                for (var i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
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
            var projectTypeGuids = string.Empty;

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
            var projectTypeGuids = GetProjectTypeGuids(project);
            if (string.IsNullOrEmpty(projectTypeGuids))
                return false;

            projectTypeGuids = StringFormatting.RemoveBracesToUpper(projectTypeGuids);
            var guids = projectTypeGuids.Split(';');

            return guids.Contains(ExtensionConstants.UnitTestProjectType.ToString(), StringComparer.InvariantCultureIgnoreCase);
        }

        public static string GetAssemblyPath(Project project)
        {
            var fullPath = project.Properties.Item(Resource.Constant_ProjectProperties_FullPath).Value.ToString();
            var outputPath = project.ConfigurationManager.ActiveConfiguration.Properties.Item(Resource.Constant_ProjectProperties_OutputPath).Value.ToString();
            var outputDir = Path.Combine(fullPath, outputPath);
            var outputFileName = project.Properties.Item(Resource.Constant_ProjectProperties_OutputFileName).Value.ToString();
            var assemblyPath = Path.Combine(outputDir, outputFileName);

            return assemblyPath;
        }

        public static bool BuildProject(Project project)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return false;

            var solutionBuild = dte.Solution.SolutionBuild;
            solutionBuild.BuildProject(dte.Solution.SolutionBuild.ActiveConfiguration.Name, project.UniqueName, true);

            //0 = no errors
            return solutionBuild.LastBuildInfo <= 0;
        }

        public static string GetProjectPath(Project project)
        {
            var path = project.FullName;

            if (File.Exists(path))
                path = Path.GetDirectoryName(path);

            if (!string.IsNullOrEmpty(path))
                return path;

            OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_ErrorGetProjectPath, MessageType.Error);
            return null;
        }

        public static string GetOutputFile(Project project)
        {
            var outputFileName = project.Properties.Item(Resource.Constant_ProjectProperties_OutputFileName).Value.ToString();
            var path = GetOutputPath(project);
            return path == null ?
                null :
                Path.Combine(path, outputFileName);
        }

        public static string GetOutputPath(Project project)
        {
            var configurationManager = project.ConfigurationManager;
            if (configurationManager == null) return null;

            var activeConfiguration = configurationManager.ActiveConfiguration;
            var outputPath = activeConfiguration.Properties.Item(Resource.Constant_ProjectProperties_OutputPath).Value.ToString();
            var absoluteOutputPath = string.Empty;
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
                var existingReference = vsproject.References.Find(referenceName);
                if (existingReference != null)
                    return;

                vsproject.References.Add(referenceName);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_ErrorAddProjectReference}: {referenceName}", ex);
            }
        }

        public static bool IsFileInProjectFile(string projectFilePath, string fileRelativePath)
        {
            try
            {
                var p = new Microsoft.Build.Evaluation.Project(projectFilePath, null, null, new ProjectCollection());

                foreach (var projectItem in p.Items)
                {
                    if (projectItem.EvaluatedInclude == fileRelativePath)
                        return true;
                }
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.Error_UnableToReadProjectFile, ex);
            }

            return false;
        }

        public static void MovePdbFile(Project project, string assemblyFilePath)
        {
            try
            {
                var pdbPath = Path.ChangeExtension(assemblyFilePath, "pdb");
                if (pdbPath == null)
                    return;

                var pdbFile = new FileInfo(pdbPath);
                if (pdbFile.Exists)
                    pdbFile.MoveTo(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + Path.GetFileName(pdbPath)));

            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.Error_ErrorMovingPDB, ex);
            }
        }
    }
}