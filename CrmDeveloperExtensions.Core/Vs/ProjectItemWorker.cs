using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace CrmDeveloperExtensions.Core.Vs
{
    public static class ProjectItemWorker
    {
        //https://www.mztools.com/articles/2014/MZ2014006.aspx



        public static void ProcessProjectItem(IVsSolution solutionService, EnvDTE.Project project)
        {
            IVsHierarchy projectHierarchy = null;

            if (solutionService.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) == VSConstants.S_OK)
            {
                if (projectHierarchy != null)
                {
                    foreach (EnvDTE.ProjectItem projectItem in project.ProjectItems)
                    {
                        string fileFullName = null;
                        uint itemId;

                        try
                        {
                            fileFullName = projectItem.get_FileNames(0);
                        }
                        catch
                        {
                        }

                        if (!string.IsNullOrEmpty(fileFullName))
                        {
                            if (projectHierarchy.ParseCanonicalName(fileFullName, out itemId) == VSConstants.S_OK)
                            {
                                MessageBox.Show("File: " + fileFullName + "\r\n" + "Item Id: 0x" + itemId.ToString("X"));
                            }
                        }
                    }

                    
                }
            }
        }
    }
}
