using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using NuGet.VisualStudio;
using System;
using System.Linq;
using TemplateWizards;
using VSLangProj;

namespace PluginDeployer
{
    public static class IlMergeHandler
    {
        public static bool Install(DTE dte, Project project)
        {
            try
            {
                var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
                if (componentModel == null)
                    return false;

                var installer = componentModel.GetService<IVsPackageInstaller>();

                NuGetProcessor.InstallPackage(installer, project,
                    CrmDeveloperExtensions2.Core.ExtensionConstants.IlMergeNuGet, null);

                return true;
            }
            catch (Exception)
            {
                //MessageBox.Show("Error installing MSBuild.ILMerge.Task" + Environment.NewLine + Environment.NewLine + ex.Message);
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

                NuGetProcessor.UnInstallPackage(uninstaller, project, CrmDeveloperExtensions2.Core.ExtensionConstants.IlMergeNuGet);

                //SetIlMergeTooltip(false);
                //_isIlMergeInstalled = false;
                return true;
            }
            catch (Exception)
            {
                //MessageBox.Show("Error uninstalling MSBuild.ILMerge.Task" + Environment.NewLine + Environment.NewLine + ex.Message);
                return false;
            }
        }

        public static void SetReferenceCopyLocal(Project project, bool copyLocal)
        {
            string[] excludedAssemblies = {
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmSdk,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MMicrosoftCrmSdkProxy,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmSdkDeployment,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmClient,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmPortal,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmSdkWorkflow,
                CrmDeveloperExtensions2.Core.ExtensionConstants.MicrosoftXrmToolingConnector
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