using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Config;
using D365DeveloperExtensions.Core.Models;
using EnvDTE;
using SolutionPackager.ViewModels;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CoreMapping = D365DeveloperExtensions.Core.Config.Mapping;

namespace SolutionPackager.Config
{
    public static class Mapping
    {
        public static SolutionPackageConfig GetSolutionPackageConfig(Project project, string profile, ObservableCollection<CrmSolution> crmSolutions)
        {
            SpklConfig spklConfig = CoreMapping.GetSpklConfigFile(project);

            List<SolutionPackageConfig> spklSolutionPackageConfigs = spklConfig.solutions;
            if (spklSolutionPackageConfigs == null)
                return null;

            SolutionPackageConfig solutionPackageConfig = profile.StartsWith(ExtensionConstants.NoProfilesText)
                ? spklSolutionPackageConfigs[0]
                : spklSolutionPackageConfigs.FirstOrDefault(p => p.profile == profile);

            return solutionPackageConfig;
        }

        public static void AddOrUpdateSpklMapping(Project project, string profile, SolutionPackageConfig solutionPackageConfig)
        {
            SpklConfig spklConfig = CoreMapping.GetSpklConfigFile(project);

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.solutions[0] = solutionPackageConfig;
            else
            {
                SolutionPackageConfig existingSolutionPackageConfig = spklConfig.solutions.FirstOrDefault(s => s.profile == profile);
                if (existingSolutionPackageConfig != null && solutionPackageConfig != null)
                {
                    existingSolutionPackageConfig.increment_on_import = solutionPackageConfig.increment_on_import;
                    existingSolutionPackageConfig.map = solutionPackageConfig.map;
                    existingSolutionPackageConfig.packagetype = solutionPackageConfig.packagetype;
                    existingSolutionPackageConfig.packagepath = solutionPackageConfig.packagepath.Replace("/", string.Empty);
                    existingSolutionPackageConfig.solution_uniquename = solutionPackageConfig.solution_uniquename;
                    existingSolutionPackageConfig.solutionpath = FormatSolutionName(solutionPackageConfig.solutionpath);
                }
            }

            string projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        private static string FormatSolutionName(string solutionName)
        {
            return string.IsNullOrEmpty(solutionName)
                ? string.Empty :
                $"{solutionName}_{{0}}_{{1}}_{{2}}_{{3}}.zip";
        }
    }
}