using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using System.Windows;
using System.Windows.Forms;
using Resource = TemplateWizards.Resources.Resource;

namespace TemplateWizards
{
    public class SdkToolsInstaller
    {
        public static void InstallNuGetCliPackage(string packageName)
        {
            var packages = NuGetRetriever.PackageLister.GetPackagesById(packageName);

            var nuGetPicker = new NuGetPicker(packageName, packages);
            var result = nuGetPicker.ShowDialog();
            if (!result.HasValue || result.Value == false)
                return;

            var selectedPackage = nuGetPicker.SelectedPackage;

            var installPath = GetInstallPath();
            if (string.IsNullOrEmpty(installPath))
                return;

            var installed = NuGetCliProcessor.Install(installPath, selectedPackage);
            if (!installed)
                return;

            if (selectedPackage.Id == ExtensionConstants.MicrosoftCrmSdkCoreTools)
                UpdateCorePath(installPath, selectedPackage);

            if (selectedPackage.Id == ExtensionConstants.MicrosoftCrmSdkXrmToolingPrt)
                UpdatePrtPath(installPath, selectedPackage);
        }

        private static void UpdatePrtPath(string installPath, NuGetPackage selectedPackage)
        {
            var prtPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.PluginRegistrationToolPath);
            if (!string.IsNullOrEmpty(prtPath))
                return;

            var updatePrt = System.Windows.MessageBox.Show(Resource.ConfirmMessage_UpdatePrtPath, Resource.ConfirmMessage_UpdateUserSetting_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (updatePrt != MessageBoxResult.Yes)
                return;

            var newPrtPath =
                $@"{installPath}\{ExtensionConstants.MicrosoftCrmSdkXrmToolingPrt.ToLower()}.{selectedPackage.VersionText}\tools";

            UserOptionsHelper.SetOption(UserOptionProperties.PluginRegistrationToolPath, newPrtPath);
        }

        private static void UpdateCorePath(string installPath, NuGetPackage selectedPackage)
        {
            var spPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.SolutionPackagerToolPath);
            var crmSvcPath = UserOptionsHelper.GetOption<string>(UserOptionProperties.CrmSvcUtilToolPath);

            if (!string.IsNullOrEmpty(spPath) && !string.IsNullOrEmpty(crmSvcPath))
                return;

            var updateCore = System.Windows.MessageBox.Show(Resource.ConfirmMessage_UpdateCrmSvcUtilSolutionPackagerPath, Resource.ConfirmMessage_UpdateUserSetting_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (updateCore != MessageBoxResult.Yes)
                return;

            var newCorePath =
                $@"{installPath}\{ExtensionConstants.MicrosoftCrmSdkCoreTools.ToLower()}.{selectedPackage.VersionText}\content\bin\coretools";

            UserOptionsHelper.SetOption(UserOptionProperties.CrmSvcUtilToolPath, newCorePath);
            UserOptionsHelper.SetOption(UserOptionProperties.SolutionPackagerToolPath, newCorePath);
        }

        private static string GetInstallPath()
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            var result = folderBrowserDialog.ShowDialog();

            return result != DialogResult.OK
                ? null
                : folderBrowserDialog.SelectedPath;
        }
    }
}