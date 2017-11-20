using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using SolutionPackager.ViewModels;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CoreMapping = CrmDeveloperExtensions2.Core.Config.Mapping;

namespace SolutionPackager.Config
{
    public static class Mapping
    {
        public static SolutionPackageConfig GetSolutionPackageConfig(Project project, string profile, ObservableCollection<CrmSolution> crmSolutions)
        {
            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            SpklConfig spklConfig = CoreMapping.GetSpklConfigFile(projectPath, project);

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
            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            SpklConfig spklConfig = CoreMapping.GetSpklConfigFile(projectPath, project);

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