using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Models;
using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class Profiles
    {
        public static List<string> GetProfiles(string projectPath, ToolWindowType toolWindowType)
        {
            SpklConfig spklConfig = ConfigFile.GetSpklConfigFile(projectPath);
            switch (toolWindowType)
            {
                case ToolWindowType.PluginDeployer:
                    return GetConfigProfiles(projectPath, spklConfig.plugins);
                case ToolWindowType.PluginTraceViewer:
                    return null;
                case ToolWindowType.SolutionPackager:
                    return GetConfigProfiles(projectPath, spklConfig.solutions);
                case ToolWindowType.WebResourceDeployer:
                    return GetConfigProfiles(projectPath, spklConfig.webresources);
                default:
                    return null;
            }
        }

        public static List<string> GetConfigProfiles<T>(string projectPath, List<T> configs)
        {
            List<string> profiles = new List<string>();

            int i = 1;
            foreach (dynamic config in configs)
            {
                profiles.Add(string.IsNullOrEmpty(config.profile)
                    ? $"{ExtensionConstants.NoProfilesText} {i}"
                    : config.profile);
                i++;
            }

            return profiles;
        }
    }
}