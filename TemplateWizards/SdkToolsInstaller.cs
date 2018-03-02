using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using Resource = TemplateWizards.Resources.Resource;

namespace TemplateWizards
{
    public class SdkToolsInstaller
    {
        public static void InstallNuGetCliPackage(string packageName)
        {
            List<NuGetPackage> packages =
                NuGetRetriever.PackageLister.GetPackagesbyId(packageName);

            NuGetPicker nuGetPicker = new NuGetPicker(packageName, packages);
            bool? result = nuGetPicker.ShowDialog();
            if (!result.HasValue || result.Value == false)
                return;

            NuGetPackage selectedPackage = nuGetPicker.SelectedPackage;

            string installPath = GetInstallPath();
            if (string.IsNullOrEmpty(installPath))
                return;

            bool installed = NuGetCliProcessor.Install(installPath, selectedPackage);
            if (!installed)
                return;

            if (selectedPackage.Id == ExtensionConstants.MicrosoftCrmSdkCoreTools)
                UpdateCorePath(installPath, selectedPackage);

            if (selectedPackage.Id == ExtensionConstants.MicrosoftCrmSdkXrmToolingPrt)
                UpdatePrtPath(installPath, selectedPackage);
        }

        private static void UpdatePrtPath(string installPath, NuGetPackage selectedPackage)
        {
            string prtPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.PluginRegistrationToolPath);
            if (!string.IsNullOrEmpty(prtPath))
                return;

            MessageBoxResult updatePrt = System.Windows.MessageBox.Show(Resource.ConfirmMessage_UpdatePrtPath, Resource.ConfirmMessage_UpdateUserSetting_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (updatePrt != MessageBoxResult.Yes)
                return;

            string newPrtPath =
                $@"{installPath}\{ExtensionConstants.MicrosoftCrmSdkXrmToolingPrt.ToLower()}.{selectedPackage.VersionText}\tools";

            UserOptionsHelper.SetOption(UserOptionProperties.PluginRegistrationToolPath, newPrtPath);
        }

        private static void UpdateCorePath(string installPath, NuGetPackage selectedPackage)
        {
            string spPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.SolutionPackagerToolPath);
            string crmSvcPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.CrmSvcUtilToolPath);

            if (!string.IsNullOrEmpty(spPath) && !string.IsNullOrEmpty(crmSvcPath))
                return;

            MessageBoxResult updateCore = System.Windows.MessageBox.Show(Resource.ConfirmMessage_UpdateCrmSvcUtilSolutionPackagerPath, Resource.ConfirmMessage_UpdateUserSetting_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (updateCore != MessageBoxResult.Yes)
                return;

            string newCorePath =
                $@"{installPath}\{ExtensionConstants.MicrosoftCrmSdkCoreTools.ToLower()}.{selectedPackage.VersionText}\content\bin\coretools";

            UserOptionsHelper.SetOption(UserOptionProperties.CrmSvcUtilToolPath, newCorePath);
            UserOptionsHelper.SetOption(UserOptionProperties.SolutionPackagerToolPath, newCorePath);
        }

        private static string GetInstallPath()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog.ShowDialog();

            return result != DialogResult.OK
                ? null
                : folderBrowserDialog.SelectedPath;
        }
    }
}