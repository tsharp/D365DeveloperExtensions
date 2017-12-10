using CrmDeveloperExtensions2.Core;
using EnvDTE;
using NuGet.VisualStudio;
using System;
using System.Windows;
using TemplateWizards.Enums;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;

namespace TemplateWizards
{
    public static class NuGetProcessor
    {


        public static void InstallPackage(IVsPackageInstaller installer, Project project, string package, string version)
        {
            try
            {
                string nuGetSource = "https://www.nuget.org/api/v2/";
                StatusBar.SetStatusBarValue(Resources.Resource.NuGetPackageInstallingStatusBarMessage + ": " + package + " " + version);
                installer.InstallPackage(nuGetSource, project, package, version, false);
            }
            catch (Exception ex)
            {

                //TODO: handle this error better if unable to connect - displays large error detail otherwise
                MessageBox.Show(Resources.Resource.NuGetPackageInstallFailureMessage + ": " + ex.Message);
            }
            finally
            {
                StatusBar.ClearStatusBarValue();
            }
        }

        public static void UnInstallPackage(IVsPackageUninstaller uninstaller, Project project, string package)
        {
            try
            {
                StatusBar.SetStatusBarValue(Resources.Resource.NuGetPackageUninstallingStatusBarMessage + ": " + package);

                uninstaller.UninstallPackage(project, package, true);
            }
            catch (Exception)
            {

                //MessageBox.Show(Resources.Resource.NuGetPackageInstallFailureMessage + ": " + ex.Message);
            }
            finally
            {
                StatusBar.ClearStatusBarValue();
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