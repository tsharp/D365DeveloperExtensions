using System;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio;

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
    }
}
