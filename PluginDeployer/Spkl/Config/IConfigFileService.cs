using System.Collections.Generic;

namespace PluginDeployer.Spkl.Config
{
    public interface IConfigFileService
    {
        List<ConfigFile> FindConfig(string folder, bool raiseErrorIfNotFound = true);
    }
}