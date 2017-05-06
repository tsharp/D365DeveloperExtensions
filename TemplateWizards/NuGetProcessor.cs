using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using CrmDeveloperExtensions.Core.Models;
using EnvDTE;
using NuGet.VisualStudio;
using TemplateWizards.Enums;

namespace TemplateWizards
{
    public static class NuGetProcessor
    {
        public static void InstallPackage(IVsPackageInstaller installer, Project project, string package, string version)
        {
            try
            {
                string nuGetSource = "https://www.nuget.org/api/v2/";
                installer.InstallPackage(nuGetSource, project, package, version, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: Error Installing NuGet Package: " + ex.Message);
            }
        }

        public static string DetermineClientType(string coreVersion)
        {
            Version version = CrmDeveloperExtensions.Core.Versioning.StringToVersion(coreVersion);
            int result = version.CompareTo(new Version(6, 1, 0));
            return result >= 0 ? "Microsoft.CrmSdk.XrmTooling.CoreAssembly" 
                               : "Microsoft.CrmSdk.Extensions";
        }

        public static int GetNextPackage(PackageValue package, bool getWorkflow, bool getClient)
        {
            switch (package)
            {
                case PackageValue.Core:
                    if (getWorkflow) return 2;
                    if (getClient) return 3;
                    return 0;
                case PackageValue.Workflow:
                    if (getClient) return 3;
                    return 0;
                default:
                    return 0;
            }
        }
    }
}