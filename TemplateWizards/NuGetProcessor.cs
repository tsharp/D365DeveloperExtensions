using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using EnvDTE;
using NLog;
using NuGet.VisualStudio;
using System;
using System.Windows;
using StatusBar = D365DeveloperExtensions.Core.StatusBar;

namespace TemplateWizards
{
    public static class NuGetProcessor
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void InstallPackage(IVsPackageInstaller installer, Project project, string package, string version)
        {
            try
            {
                StatusBar.SetStatusBarValue(Resources.Resource.NuGetPackageInstallingStatusBarMessage + ": " + package + " " + version);
                installer.InstallPackage(ExtensionConstants.NuGetSourceUrl, project, package, version, false);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.NuGetPackageInstallFailureMessage, ex);
                MessageBox.Show(Resources.Resource.NuGetPackageInstallFailureMessage);
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
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.NuGetPackageInstallFailureMessage, ex);
                MessageBox.Show(Resources.Resource.NuGetPackageInstallFailureMessage + ": " + ex.Message);
            }
            finally
            {
                StatusBar.ClearStatusBarValue();
            }
        }

        public static string DetermineClientType(string coreVersion)
        {
            var version = Versioning.StringToVersion(coreVersion);
            var result = version.CompareTo(new Version(6, 1, 0));
            return result >= 0 ? Resources.Resource.SdkAssemblyXrmTooling
                               : Resources.Resource.SdkAssemblyExtensions;
        }

        public static int GetNextPackage(TemplatePackageType templatePackage, bool getWorkflow, bool getClient)
        {
            switch (templatePackage)
            {
                case TemplatePackageType.Core:
                    if (getWorkflow)
                        return 2;
                    if (getClient)
                        return 3;
                    return 0;
                case TemplatePackageType.Workflow:
                    if (getClient)
                        return 3;
                    return 0;
                default:
                    return 0;
            }
        }
    }
}