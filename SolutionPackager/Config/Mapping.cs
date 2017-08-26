using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using System;
using System.Collections.ObjectModel;
using SolutionPackager.ViewModels;
using CoreMapping = CrmDeveloperExtensions2.Core.Config.Mapping;

namespace SolutionPackager.Config
{
    public static class Mapping
    {
        public static CrmDevExSolutionPackage HandleMappings(string solutionPath, Project project, ObservableCollection<CrmSolution> crmSolutions, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CoreMapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CoreMapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            if (crmDevExConfigOrgMap.SolutionPackage == null)
                return null;

            foreach (CrmSolution crmSolution in crmSolutions)
            {
                if (crmSolution.SolutionId == crmDevExConfigOrgMap.SolutionPackage.SolutionId)
                    return crmDevExConfigOrgMap.SolutionPackage;
            }

            crmDevExConfigOrgMap.SolutionPackage = null;

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);

            return null;
        }

        public static void AddOrUpdateMapping(string solutionPath, Project project, CrmDevExSolutionPackage solutionPackage, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CoreMapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CoreMapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            crmDevExConfigOrgMap.SolutionPackage = solutionPackage;

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }
    }
}