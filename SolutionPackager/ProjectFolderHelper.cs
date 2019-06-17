using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
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
            var projectItem = e.ProjectItem;
            var itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            var newItemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]).TrimEnd('/');
            projectFolders.Add(newItemName);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }

        public static ObservableCollection<string> FolderRemoved(ProjectItemRemovedEventArgs e, ObservableCollection<string> projectFolders)
        {
            var projectItem = e.ProjectItem;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            var itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var itemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            projectFolders.Remove(itemName);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }

        public static ObservableCollection<string> FolderRenamed(ProjectItemRenamedEventArgs e, ObservableCollection<string> projectFolders)
        {
            var projectItem = e.ProjectItem;
            if (projectItem.Name == null)
                return projectFolders;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return projectFolders;

            var oldName = e.OldName;
            var itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return projectFolders;

            var newItemPath = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            var index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
            if (index == -1)
                return projectFolders;

            var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

            projectFolders.Remove(oldItemPath);

            return new ObservableCollection<string>(projectFolders.OrderBy(s => s));
        }
    }
}
