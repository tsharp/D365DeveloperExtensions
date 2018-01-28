using D365DeveloperExtensions.Core.Models;
using EnvDTE;

namespace D365DeveloperExtensions.Core.Config
{
    public static class Mapping
    {
        public static SpklConfig GetSpklConfigFile(Project project)
        {
            string projectPath = Vs.ProjectWorker.GetProjectPath(project);

            if (ConfigFile.SpklConfigFileExists(projectPath))
                return ConfigFile.GetSpklConfigFile(project);

            ConfigFile.CreateSpklConfigFile(project);
            return ConfigFile.GetSpklConfigFile(project);
        }
    }
}