using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class Mapping
    {
        public static SpklConfig GetSpklConfigFile(string projectPath, Project project)
        {
            if (ConfigFile.SpklConfigFileExists(projectPath))
                return ConfigFile.GetSpklConfigFile(projectPath);

            ConfigFile.CreateSpklConfigFile(project);
            return ConfigFile.GetSpklConfigFile(Vs.ProjectWorker.GetProjectPath(project));
        }
    }
}