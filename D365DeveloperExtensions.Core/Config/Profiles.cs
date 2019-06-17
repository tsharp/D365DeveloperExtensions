using D365DeveloperExtensions.Core.Enums;
using EnvDTE;
using System.Collections.Generic;

namespace D365DeveloperExtensions.Core.Config
{
    public static class Profiles
    {
        /// <summary>Gets profiles from the configuration file.</summary>
        /// <param name="project">The project.</param>
        /// <param name="toolWindowType">Type of the tool window.</param>
        /// <returns>List of profiles.</returns>
        public static List<string> GetProfiles(Project project, ToolWindowType toolWindowType)
        {
            if (toolWindowType == ToolWindowType.PluginTraceViewer)
                return null;

            var spklConfig = ConfigFile.GetSpklConfigFile(project);

            switch (toolWindowType)
            {
                case ToolWindowType.PluginDeployer:
                    return GetConfigProfiles(spklConfig.plugins);
                case ToolWindowType.SolutionPackager:
                    return GetConfigProfiles(spklConfig.solutions);
                case ToolWindowType.WebResourceDeployer:
                    return GetConfigProfiles(spklConfig.webresources);
                default:
                    return null;
            }
        }

        private static List<string> GetConfigProfiles<T>(List<T> configs)
        {
            if (configs == null)
                return new List<string>();

            var profiles = new List<string>();

            var i = 1;
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