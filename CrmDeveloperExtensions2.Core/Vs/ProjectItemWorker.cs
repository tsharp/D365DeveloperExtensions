using System;
using System.Windows;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions2.Core.Vs
{
    public static class ProjectItemWorker
    {
        //https://www.mztools.com/articles/2014/MZ2014006.aspx

        public static void ProcessProjectItem(IVsSolution solutionService, EnvDTE.Project project)
        {
            IVsHierarchy projectHierarchy;

            if (solutionService.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) !=
                VSConstants.S_OK) return;

            if (projectHierarchy == null)
                return;

            foreach (EnvDTE.ProjectItem projectItem in project.ProjectItems)
            {
                string fileFullName = null;
                uint itemId;

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

                if (projectHierarchy.ParseCanonicalName(fileFullName, out itemId) == VSConstants.S_OK)
                    MessageBox.Show("File: " + fileFullName + "\r\n" + "Item Id: 0x" + itemId.ToString("X"));
            }
        }

        public static uint GetProjectItemId(IVsSolution solutionService, string projectUniqueName, EnvDTE.ProjectItem projectItem)
        {
            IVsHierarchy projectHierarchy;

            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out projectHierarchy) != VSConstants.S_OK)
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

            uint itemId;
            return projectHierarchy.ParseCanonicalName(fileFullName, out itemId) == VSConstants.S_OK 
                ? itemId 
                : UInt32.MinValue;
        }

        public static EnvDTE.ProjectItem GetProjectItemFromItemId(IVsSolution solutionService, string projectUniqueName, uint projectItemId)
        {
            IVsHierarchy projectHierarchy;

            if (solutionService.GetProjectOfUniqueName(projectUniqueName, out projectHierarchy) != VSConstants.S_OK)
                return null;

            if (projectHierarchy == null)
                return null;

            object objProjectItem;
            projectHierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out objProjectItem);

            var projectItem = objProjectItem as EnvDTE.ProjectItem;

            return projectItem;
        }
    }
}
