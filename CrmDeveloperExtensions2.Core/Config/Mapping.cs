using CrmDeveloperExtensions2.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class Mapping
    {
        public static void UpdateProjectName(string solutionPath, string oldProjectUniqueName, string newProjectUniqueName)
        {
            CrmDexExConfig crmDexExConfig = ConfigFile.GetConfigFile(solutionPath);
            if (crmDexExConfig == null)
                return;

            bool updated = false;
            foreach (CrmDevExConfigOrgMap crmDevExConfigOrgMap in crmDexExConfig.CrmDevExConfigOrgMaps)
                if (crmDevExConfigOrgMap.ProjectUniqueName.Equals(oldProjectUniqueName, StringComparison.InvariantCultureIgnoreCase))
                {
                    crmDevExConfigOrgMap.ProjectUniqueName = newProjectUniqueName;
                    updated = true;
                }

            if (updated)
                ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        public static CrmDevExConfigOrgMap GetOrgMap(ref CrmDexExConfig crmDexExConfig, Guid organizationId, string projectUniqueName)
        {
            CrmDevExConfigOrgMap orgMap = crmDexExConfig.CrmDevExConfigOrgMaps.FirstOrDefault(o => o.OrganizationId == organizationId);
            if (orgMap != null)
                return orgMap;

            orgMap = new CrmDevExConfigOrgMap
            {
                OrganizationId = organizationId,
                ProjectUniqueName = projectUniqueName,
                WebResources = new List<CrmDexExConfigWebResource>(),
                SolutionPackage = null,
                Assembly = null
            };

            crmDexExConfig.CrmDevExConfigOrgMaps.Add(orgMap);

            return orgMap;
        }

        public static CrmDexExConfig GetConfigFile(string solutionPath, string projectUniqueName, Guid organizationId)
        {
            return !ConfigFile.ConfigFileExists(solutionPath) ?
                ConfigFile.CreateConfigFile(organizationId, projectUniqueName, solutionPath) :
                ConfigFile.GetConfigFile(solutionPath);
        }

        public static void RemoveProject(string solutionPath, string projectUniqueName)
        {
            if (!ConfigFile.ConfigFileExists(solutionPath))
                return;

            CrmDexExConfig crmDexExConfig = ConfigFile.GetConfigFile(solutionPath);
            var orgMapsToDelete = crmDexExConfig.CrmDevExConfigOrgMaps.Where(o => o.ProjectUniqueName == projectUniqueName);

            if (orgMapsToDelete.Any())
            {
                MessageBoxResult result = MessageBox.Show("Delete project mappings?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
                if (result != MessageBoxResult.Yes)
                    return;
            }

            crmDexExConfig.CrmDevExConfigOrgMaps.RemoveAll(o => o.ProjectUniqueName == projectUniqueName);

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }
    }
}