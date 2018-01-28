using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using NLog;
using NuGet.VisualStudio;
using PluginDeployer.Resources;
using System;
using System.Linq;
using System.Windows;
using TemplateWizards;
using VSLangProj;

namespace PluginDeployer
{
    public static class IlMergeHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool TogleIlMerge(DTE dte, Project selecteProject, bool isIlMergeInstalled)
        {
            if (!isIlMergeInstalled)
            {
                bool installed = Install(dte, selecteProject);

                // Set CRM Assemblies to "Copy Local = false" to prevent merging
                if (installed)
                {
                    SetReferenceCopyLocal(selecteProject, false);
                    return true;
                }

                MessageBox.Show(Resource.ErrorMessage_ErrorInstallingILMerge);
                isIlMergeInstalled = true;
            }
            else
            {
                bool uninstalled = Uninstall(dte, selecteProject);

                // Reset CRM Assemblies to "Copy Local = true"
                if (uninstalled)
                {
                    SetReferenceCopyLocal(selecteProject, true);
                    return false;
                }

                MessageBox.Show(Resource.ErrorMessage_ErrorUninstallingILMerge);
                isIlMergeInstalled = false;
            }

            return isIlMergeInstalled;
        }

        public static bool Install(DTE dte, Project project)
        {
            try
            {
                var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
                if (componentModel == null)
                    return false;

                var installer = componentModel.GetService<IVsPackageInstaller>();

                NuGetProcessor.InstallPackage(installer, project, ExtensionConstants.IlMergeNuGet, null);

                OutputLogger.WriteToOutputWindow(Resource.Message_ILMergeInstalled, MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorInstallingILMerge, ex);

                return false;
            }
        }

        public static bool Uninstall(DTE dte, Project project)
        {
            try
            {
                var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
                if (componentModel == null)
                    return false;

                var uninstaller = componentModel.GetService<IVsPackageUninstaller>();

                NuGetProcessor.UnInstallPackage(uninstaller, project, ExtensionConstants.IlMergeNuGet);

                OutputLogger.WriteToOutputWindow(Resource.Message_ILMergeUninstalled, MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorUninstallingILMerge, ex);

                return false;
            }
        }

        public static void SetReferenceCopyLocal(Project project, bool copyLocal)
        {
            string[] excludedAssemblies = {
                ExtensionConstants.MicrosoftXrmSdk,
                ExtensionConstants.MMicrosoftCrmSdkProxy,
                ExtensionConstants.MicrosoftXrmSdkDeployment,
                ExtensionConstants.MicrosoftXrmClient,
                ExtensionConstants.MicrosoftXrmPortal,
                ExtensionConstants.MicrosoftXrmSdkWorkflow,
                ExtensionConstants.MicrosoftXrmToolingConnector
            };

            if (!(project.Object is VSProject vsproject))
                return;

            foreach (Reference reference in vsproject.References)
            {
                if (reference.SourceProject != null) continue;

                if (excludedAssemblies.Contains(reference.Name))
                    reference.CopyLocal = copyLocal;
            }
        }
    }
}