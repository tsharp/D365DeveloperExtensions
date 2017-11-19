using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class Mapping
    {
        public static SpklConfig GetSpklConfigFile(string projectPath, Project project)
        {
            return !ConfigFile.SpklConfigFileExists(projectPath) ?
                ConfigFile.CreateSpklConfigFile(project) :
                ConfigFile.GetSpklConfigFile(projectPath);
        }
    }
}