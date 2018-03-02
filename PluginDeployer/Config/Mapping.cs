using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using EnvDTE;
using System.Collections.Generic;
using System.Linq;

namespace PluginDeployer.Config
{
    public static class Mapping
    {
        public static PluginDeployConfig GetSpklPluginConfig(Project project, string profile)
        {
            SpklConfig spklConfig = D365DeveloperExtensions.Core.Config.Mapping.GetSpklConfigFile(project);

            List<PluginDeployConfig> spklPluginDeployConfigs = spklConfig.plugins;
            if (spklPluginDeployConfigs == null)
                return null;

            return profile.StartsWith(ExtensionConstants.NoProfilesText)
                ? spklPluginDeployConfigs[0]
                : spklPluginDeployConfigs.FirstOrDefault(p => p.profile == profile);
        }
    }
}