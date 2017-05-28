using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Controls;

namespace CrmDeveloperExtensions.Core.Vs
{
    public static class ProjectWorker
    {
        static readonly string[] Extensions = { "HTM", "HTML", "CSS", "JS", "XML", "PNG", "JPG", "GIF", "XAP", "XSL", "XSLT", "ICO", "TS" };
        static readonly string[] FolderExtensions = { "BUNDLE", "TT" };

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

        public static Project GetProjectByName(string projectName)
        {
            IList<Project> projects = GetProjects();
            foreach (Project project in projects)
            {
                if (project.Name != projectName) continue;

                return project;
            }

            return null;
        }

        private static IList<Project> GetProjects()
        {
            List<Project> list = new List<Project>();

            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (dte == null)
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

            return list;
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

        public static ObservableCollection<MenuItem> GetProjectFoldersForMenu(string projectName)
        {
            List<string> projectFolders = new List<string>();
            ObservableCollection<MenuItem> projectMenuItems = new ObservableCollection<MenuItem>();
            Project project = GetProjectByName(projectName);
            if (project == null)
                return projectMenuItems;

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

            foreach (string projectFolder in projectFolders)
            {
                MenuItem item = new MenuItem
                {
                    Header = projectFolder
                };

                projectMenuItems.Add(item);
            }

            return projectMenuItems;
        }

        private static ObservableCollection<string> GetFolders(ProjectItem projectItem, string path)
        {
            ObservableCollection<string> projectFolders = new ObservableCollection<string>();
            if (new Guid(projectItem.Kind) == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                projectFolders.Add(path + "/" + projectItem.Name);
                for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    var folders = GetFolders(projectItem.ProjectItems.Item(i), path + "/" + projectItem.Name);
                    foreach (string folder in folders)
                        projectFolders.Add(folder);
                }
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
    }
}
