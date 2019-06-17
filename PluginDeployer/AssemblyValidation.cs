using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using PluginDeployer.Resources;
using PluginDeployer.Spkl;
using System;
using System.IO;
using System.Linq;
using System.Windows;

namespace PluginDeployer
{
    public static class AssemblyValidation
    {
        public static bool ValidatePluginDeployConfig(PluginDeployConfig pluginDeployConfig)
        {
            if (pluginDeployConfig == null)
            {
                MessageBox.Show($"{Resource.MessageBox_MissingPluginsSpklConfig}: {ExtensionConstants.SpklConfigFile}");
                return false;
            }

            if (string.IsNullOrEmpty(pluginDeployConfig.assemblypath))
            {
                MessageBox.Show($"{Resource.MessageBox_MissingAssemblyPathSpklConfig}: {ExtensionConstants.SpklConfigFile}");
                return false;
            }

            return true;
        }

        public static bool ValidateAssemblyPath(string assemblyPath)
        {
            if (string.IsNullOrEmpty(assemblyPath))
            {
                MessageBox.Show(Resource.Message_AssemblyPathEmpty);
                return false;
            }

            if (!Directory.Exists(assemblyPath))
            {
                MessageBox.Show($"{Resource.MessageBox_ErrorLocatingAssemblyPath}: {assemblyPath}");
                return false;
            }

            return true;
        }

        public static bool ValidateRegistrationDetails(string assemblyPath, bool isWorkflow)
        {
            var hasRegistration = RegistrationDetailsPresent(assemblyPath, isWorkflow);
            if (hasRegistration)
                return true;

            MessageBox.Show(Resource.Message_NoRegistrationDetails);

            return false;
        }

        public static bool ValidateAssemblyVersion(CrmServiceClient client, Entity foundAssembly, string projectAssemblyName, Version projectAssemblyVersion)
        {
            var serverVersion = Versioning.StringToVersion(foundAssembly.GetAttributeValue<string>("version"));
            var versionMatch = Versioning.DoAssemblyVersionsMatch(projectAssemblyVersion, serverVersion);

            if (!versionMatch)
            {
                MessageBox.Show(Resource.Message_UpdatingMajorMinorRequireRedeployment);
                return false;
            }

            if (!projectAssemblyName.Equals(foundAssembly.GetAttributeValue<string>("name"), StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show(Resource.Message_UpdatingNameRequireRedeployment);
                return false;
            }

            return true;
        }

        private static bool RegistrationDetailsPresent(string assemblyPath, bool isWorkflow)
        {
            AssemblyContainer container = null;
            try
            {
                var assemblyBytes = File.ReadAllBytes(assemblyPath);

                var assemblyFolderPath = Path.GetDirectoryName(assemblyPath);

                container = AssemblyContainer.LoadAssembly(assemblyBytes, isWorkflow, assemblyFolderPath);

                return container.PluginDatas.First().CrmPluginRegistrationAttributes.Count > 0;
            }
            finally
            {
                container?.Unload();
            }
        }
    }
}