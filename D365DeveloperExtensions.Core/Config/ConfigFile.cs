using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using D365DeveloperExtensions.Core.Vs;
using EnvDTE;
using Newtonsoft.Json;
using NLog;
using System;
using System.IO;
using System.Windows;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace D365DeveloperExtensions.Core.Config
{
    public static class ConfigFile
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>  Checks if the configuration file exists.</summary>
        /// <param name="projectPath">The project path.</param>
        /// <returns><c>true</c> if the configuration file exists, <c>false</c> otherwise.</returns>
        public static bool SpklConfigFileExists(string projectPath)
        {
            var path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

            return File.Exists(path);
        }

        /// <summary>Gets the existing or creates a new configuration file.</summary>
        /// <param name="project">The project.</param>
        /// <param name="isRetry">if set to <c>true</c> [is retry].</param>
        /// <returns>Configuration file.</returns>
        public static SpklConfig GetSpklConfigFile(Project project, bool isRetry = false)
        {
            var projectPath = ProjectWorker.GetProjectPath(project);
            var path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

            ExLogger.LogToFile(Logger, $"{Resource.Message_ReadingFile}: {path}", LogLevel.Info);

            try
            {
                SpklConfig spklConfig;
                using (var file = File.OpenText(path))
                {
                    var serializer = new JsonSerializer();
                    spklConfig = (SpklConfig)serializer.Deserialize(file, typeof(SpklConfig));
                }

                return spklConfig;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_UnableReadDeserializeConfig}: {path}", ex);
                MessageBox.Show($"{Resource.ErrorMessage_UnableReadDeserializeConfig}: {path}");

                if (!isRetry)
                    return RecreateConfig(project, path);

                throw;
            }
        }

        /// <summary>Creates a new configuration file.</summary>
        /// <param name="project">The project.</param>
        public static void CreateSpklConfigFile(Project project)
        {
            ExLogger.LogToFile(Logger, $"{Resource.Message_CreatingFile}: {ExtensionConstants.SpklConfigFile}", LogLevel.Info);

            TemplateHandler.AddFileFromTemplate(project, "CSharpSpklConfig\\CSharpSpklConfig", ExtensionConstants.SpklConfigFile);
        }

        /// <summary>Updates an configuration file.</summary>
        /// <param name="projectPath">The project path.</param>
        /// <param name="spklConfig">The configuration file.</param>
        public static void UpdateSpklConfigFile(string projectPath, SpklConfig spklConfig)
        {
            var text = JsonConvert.SerializeObject(spklConfig, Formatting.Indented);

            WriteSpklConfigFile(projectPath, text);
        }

        private static SpklConfig RecreateConfig(Project project, string configPath)
        {
            try
            {
                var fileExists = FileSystem.DoesFileExist(new[] { configPath }, true);
                if (fileExists)
                    FileSystem.RenameFile(configPath);

                CreateSpklConfigFile(project);

                var recreateConfig = GetSpklConfigFile(project, true);

                MessageBox.Show(Resource.MessageBox_RecreatedConfig);

                return recreateConfig;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_RecreateConfigFile, ex);
                throw;
            }
        }

        private static void WriteSpklConfigFile(string projectPath, string text)
        {
            var path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

            ExLogger.LogToFile(Logger, $"{Resource.Message_UpdatingFile}: {path}", LogLevel.Info);

            try
            {
                File.WriteAllText(path, text);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_UnableWriteConfig}: {path}", ex);
                MessageBox.Show($"{Resource.ErrorMessage_UnableWriteConfig}: {path}");
            }
        }
    }
}