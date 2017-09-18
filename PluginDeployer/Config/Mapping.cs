using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using PluginDeployer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginDeployer.Config
{
    public static class Mapping
    {
        public static CrmDevExAssembly HandleMappings(string solutionPath, Project project, Guid organizationId, ObservableCollection<CrmSolution> solutions, ObservableCollection<CrmAssembly> assemblies)
        {
            CrmDexExConfig crmDexExConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            CrmDevExAssembly crmDevExAssembly = crmDevExConfigOrgMap.Assembly;
            if (crmDevExAssembly == null)
                return null;

            if (assemblies.Count(a => a.AssemblyId == crmDevExAssembly.AssemblyId) < 1)
            {
                crmDevExConfigOrgMap.Assembly = null;

                ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);

                return null;
            }

            if (solutions.Count(s => s.SolutionId == crmDevExAssembly.SolutionId) >= 1)
                return crmDevExAssembly;

            crmDevExAssembly.SolutionId = solutions[0].SolutionId; //Default

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);

            return crmDevExAssembly;
        }

        public static void AddOrUpdateMapping(string solutionPath, Project project, Guid assemblyId, Guid solutionId, int deploymentType, bool backupFiles, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            if (assemblyId != Guid.Empty)
            {
                CrmDevExAssembly crmDevExAssembly = new CrmDevExAssembly
                {
                    AssemblyId = assemblyId,
                    SolutionId = solutionId,
                    DeploymentType = deploymentType,
                    BackupFiles = backupFiles
                };
                crmDevExConfigOrgMap.Assembly = crmDevExAssembly;
            }
            else
                crmDevExConfigOrgMap.Assembly = null;

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }
    }
}