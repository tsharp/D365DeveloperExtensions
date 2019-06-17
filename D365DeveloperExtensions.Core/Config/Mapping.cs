using D365DeveloperExtensions.Core.Models;
using EnvDTE;

namespace D365DeveloperExtensions.Core.Config
{
    public static class Mapping
    {
        /// <summary>Gets the existing or creates a new configuration file.</summary>
        /// <param name="project">The project.</param>
        /// <returns>Configuration file.</returns>
        public static SpklConfig GetSpklConfigFile(Project project)
        {
            var projectPath = Vs.ProjectWorker.GetProjectPath(project);

            if (ConfigFile.SpklConfigFileExists(projectPath))
                return ConfigFile.GetSpklConfigFile(project);

            ConfigFile.CreateSpklConfigFile(project);
            return ConfigFile.GetSpklConfigFile(project);
        }
    }
}