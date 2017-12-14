using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Connection;
using EnvDTE;
using Microsoft.VisualStudio;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace SolutionPackager
{
    public class ProjectFolderHelper
    {
        public static ObservableCollection<string> FolderAdded(ProjectItemAddedEventArgs e, ObservableCollection<string> projectFolders)
        {
            ProjectItem projectItem = e.ProjectItem;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            string newItemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]).TrimEnd('/');
            projectFolders.Add(newItemName);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }

        public static ObservableCollection<string> FolderRemoved(ProjectItemRemovedEventArgs e, ObservableCollection<string> projectFolders)
        {
            ProjectItem projectItem = e.ProjectItem;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var itemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            projectFolders.Remove(itemName);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }

        public static ObservableCollection<string> FolderRenamed(ProjectItemRenamedEventArgs e, ObservableCollection<string> projectFolders)
        {
            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.Name == null)
                return projectFolders;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            string oldName = e.OldName;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var newItemPath = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            int index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
            if (index == -1)
                return projectFolders;

            var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

            projectFolders.Remove(oldItemPath);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }
    }
}
