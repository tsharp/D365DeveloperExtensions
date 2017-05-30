using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using CrmDeveloperExtensions2.Core;
using EnvDTE;
using NuGet.VisualStudio;
using TemplateWizards.Enums;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;

namespace TemplateWizards
{
    public static class NuGetProcessor
    {
        public static void InstallPackage(DTE dte, IVsPackageInstaller installer, Project project, string package, string version)
        {
            try
            {
                string nuGetSource = "https://www.nuget.org/api/v2/";
                StatusBar.SetStatusBarValue(dte,
                    Resources.Resource.NuGetPackageInstallingStatusBarMessage + ": " + package + " " + version);
                installer.InstallPackage(nuGetSource, project, package, version, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Resources.Resource.NuGetPackageInstallFailureMessage + ": " + ex.Message);
            }
            finally
            {
                StatusBar.ClearStatusBarValue(dte);
            }
        }

        public static string DetermineClientType(string coreVersion)
        {
            Version version = Versioning.StringToVersion(coreVersion);
            int result = version.CompareTo(new Version(6, 1, 0));
            return result >= 0 ? Resources.Resource.SdkAssemblyXrmTooling
                               : Resources.Resource.SdkAssemblyExtensions;
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