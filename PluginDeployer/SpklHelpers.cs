using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using PluginDeployer.Spkl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace PluginDeployer
{
    public class SpklHelpers
    {
        public static bool ValidateRegistraionDetails(string assemblyPath, bool isWorkflow)
        {
            bool hasRegistrion = RegistrationDetailsPresent(assemblyPath, isWorkflow);
            if (hasRegistrion)
                return true;

            MessageBox.Show("You haven't addedd any registration details to the assembly class.");
            return false;
        }

        public static bool ValidateAssemblyVersion(CrmServiceClient client, Entity foundAssembly, string projectAssemblyName, Version projectAssemblyVersion)
        {
            Version serverVersion = Versioning.StringToVersion(foundAssembly.GetAttributeValue<string>("version"));
            bool versionMatch = Versioning.DoAssemblyVersionsMatch(projectAssemblyVersion, serverVersion);
            if (!versionMatch)
            {
                MessageBox.Show("Error Updating Assembly In CRM: Changes To Major & Minor Versions Require Redeployment");
                return false;
            }

            if (!projectAssemblyName.Equals(foundAssembly.GetAttributeValue<string>("name"), StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show("Error Updating Assembly In CRM: Changes To Assembly Name Require Redeployment");
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

                container = AssemblyContainer.LoadAssembly(assemblyBytes, isWorkflow, true);

                return container.PluginDatas.First().CrmPluginRegistrationAttributes.Count > 0;
            }
            finally
            {
                container?.Unload();
            }
        }

        public static string[] AssemblyProperties(string assemblyPath)
        {
            AssemblyContainer assemblyContainer = null;
            try
            {
                assemblyContainer = AssemblyContainer.LoadAssembly(File.ReadAllBytes(assemblyPath), false, true);

                List<PluginData> pluginDatas = assemblyContainer.PluginDatas;

                AssemblyName assemblyName = pluginDatas.First().AssemblyName;

                var assemblyProperties = assemblyName.FullName.Split(",= ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                return assemblyProperties;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        public static bool ValidateAssemblyPath(string assemblyPath)
        {
            if (string.IsNullOrEmpty(assemblyPath))
            {
                MessageBox.Show("Error assembly path is empty");
                return false;
            }

            if (!Directory.Exists(assemblyPath))
            {
                MessageBox.Show("Error locating assembly path");
                return false;
            }

            return true;
        }
    }
}