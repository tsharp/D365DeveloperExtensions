using System;
using System.Collections.Generic;
using System.IO;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace CrmDeveloperExtensions.Core.Vs
{
    public static class ProjectWorker
    {
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
            if (path != null)
            {
                var dirName = new DirectoryInfo(path).Name;
                var fileName = new FileInfo(projectFullName).Name;
                string folderProjectFileName = dirName + "\\" + fileName;

                return folderProjectFileName;
            }

            return null;
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
            DTE dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            Projects projects = dte.Solution.Projects;
            List<Project> list = new List<Project>();
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
    }
}
