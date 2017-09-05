using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Vs
{
    public static class ProjectItemWorker
    {
        private static readonly IEnumerable<string> FileKinds = new[] { VSConstants.GUID_ItemType_PhysicalFile.ToString() };
        private static readonly IEnumerable<string> FolderKinds = new[] { VSConstants.GUID_ItemType_PhysicalFolder.ToString() };
        private static readonly char[] PathSeparatorChars = { Path.DirectorySeparatorChar };
        private static readonly Dictionary<string, string> KnownNestedFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            { "web.debug.config", "web.config" },
            { "web.release.config", "web.config" }
        };

        public static void ProcessProjectItem(IVsSolution solutionService, Project project)
        {
            //https://www.mztools.com/articles/2014/MZ2014006.aspx
            if (solutionService.GetProjectOfUniqueName(project.UniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return;

            if (projectHierarchy == null)
                return;

            foreach (ProjectItem projectItem in project.ProjectItems)
            {
                string fileFullName = null;

                try
                {
                    fileFullName = projectItem.FileNames[0];
                }
                catch
                {
                    // ignored
                }

                if (string.IsNullOrEmpty(fileFullName))
                    continue;

                if (projectHierarchy.ParseCanonicalName(fileFullName, out var itemId) == VSConstants.S_OK)
                    MessageBox.Show("File: " + fileFullName + "\r\n" + "Item Id: 0x" + itemId.ToString("X"));
            }
        }

        public static uint GetProjectItemId(IVsSolution solutionService, string projectUniqueName, ProjectItem projectItem)
        {
            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return UInt32.MinValue;

            if (projectHierarchy == null)
                return UInt32.MinValue;

            string fileFullName;

            try
            {
                fileFullName = projectItem.FileNames[0];
            }
            catch
            {
                return UInt32.MinValue;
            }

            if (string.IsNullOrEmpty(fileFullName))
                return UInt32.MinValue;

            return projectHierarchy.ParseCanonicalName(fileFullName, out var itemId) == VSConstants.S_OK
                ? itemId
                : UInt32.MinValue;
        }

        public static ProjectItem GetProjectItemFromItemId(IVsSolution solutionService, string projectUniqueName, uint projectItemId)
        {
            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out var projectHierarchy) != VSConstants.S_OK)
                return null;

            if (projectHierarchy == null)
                return null;

            projectHierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out var objProjectItem);

            var projectItem = objProjectItem as ProjectItem;

            return projectItem;
        }

        public static string CreateValidFolderName(string name)
        {
            string[] illegal = { "/", "?", ":", "&", "\\", "*", "\"", "<", ">", "|", "#", "_" };
            string rxString = string.Join("|", illegal.Select(Regex.Escape));
            name = Regex.Replace(name, rxString, string.Empty);

            return name;
        }

        public static ProjectItem GetProjectItem(Project project, string path)
        {
            string folderPath = Path.GetDirectoryName(path);
            string itemName = Path.GetFileName(path);

            ProjectItems container = GetProjectItems(project, folderPath);

            if (container == null || !container.TryGetFile(itemName, out var projectItem) && !container.TryGetFolder(itemName, out projectItem))
                return null;

            return projectItem;
        }

        private static ProjectItem GetProjectItem(ProjectItems projectItems, string name, IEnumerable<string> allowedItemKinds)
        {
            try
            {
                ProjectItem projectItem = projectItems.Item(name);
                if (projectItem != null && allowedItemKinds.Contains(projectItem.Kind, StringComparer.OrdinalIgnoreCase))
                    return projectItem;
            }
            catch
            {
                //ignored
            }

            return null;
        }

        public static ProjectItems GetProjectItems(Project project, string folderPath, bool createIfNotExists = false)
        {
            if (String.IsNullOrEmpty(folderPath))
                return project.ProjectItems;

            string[] pathParts = folderPath.Split(PathSeparatorChars, StringSplitOptions.RemoveEmptyEntries);

            object cursor = project;

            string fullPath = project.GetFullPath();
            string folderRelativePath = String.Empty;

            foreach (string part in pathParts)
            {
                fullPath = Path.Combine(fullPath, part);
                folderRelativePath = Path.Combine(folderRelativePath, part);

                cursor = GetOrCreateFolder(cursor, fullPath, part, createIfNotExists);
                if (cursor == null)
                    return null;
            }

            return GetProjectItems(cursor);
        }

        private static ProjectItems GetProjectItems(object parent)
        {
            if (parent is Project project)
                return project.ProjectItems;

            if (parent is ProjectItem projectItem)
                return projectItem.ProjectItems;

            return null;
        }

        public static bool TryGetFolder(this ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            projectItem = GetProjectItem(projectItems, name, FolderKinds);

            return projectItem != null;
        }

        public static bool TryGetFile(this ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            projectItem = GetProjectItem(projectItems, name, FileKinds);

            if (projectItem == null)
                return TryGetNestedFile(projectItems, name, out projectItem);

            return projectItem != null;
        }

        private static bool TryGetNestedFile(ProjectItems projectItems, string name, out ProjectItem projectItem)
        {
            if (!KnownNestedFiles.TryGetValue(name, out var parentFileName))
                parentFileName = Path.GetFileNameWithoutExtension(name);

            ProjectItem parentProjectItem = GetProjectItem(projectItems, parentFileName, FileKinds);

            projectItem = parentProjectItem != null ?
                GetProjectItem(parentProjectItem.ProjectItems, name, FileKinds) :
                null;

            return projectItem != null;
        }

        public static string GetFullPath(this Project project)
        {
            string fullPath = project.GetPropertyValue<string>("FullPath");
            if (String.IsNullOrEmpty(fullPath))
                return fullPath;

            if (File.Exists(fullPath))
                fullPath = Path.GetDirectoryName(fullPath);

            return fullPath;
        }

        public static T GetPropertyValue<T>(this Project project, string propertyName)
        {
            if (project.Properties == null)
                return default(T);

            try
            {
                Property property = project.Properties.Item(propertyName);
                if (property != null)
                    return (T)property.Value;
            }
            catch (ArgumentException)
            {
                //ignored
            }

            return default(T);
        }

        // 'parentItem' can be either a Project or ProjectItem
        private static ProjectItem GetOrCreateFolder(object parentItem, string fullPath, string folderName, bool createIfNotExists)
        {
            if (parentItem == null)
                return null;

            ProjectItems projectItems = GetProjectItems(parentItem);
            if (projectItems.TryGetFolder(folderName, out var subFolder))
                return subFolder;

            if (!createIfNotExists)
                return null;

            try
            {
                return projectItems.AddFromDirectory(fullPath);
            }
            catch (NotImplementedException)
            {
                // This is the case for F#'s project system, we can't add from directory so we fall back to this impl
                return projectItems.AddFolder(folderName);
            }
        }
    }
}
