using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using D365DeveloperExtensions.Core.Vs;
using EnvDTE;
using Newtonsoft.Json;
using NLog;
using System;
using System.IO;
using System.Windows;

namespace D365DeveloperExtensions.Core.Config
{
    public static class ConfigFile
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool SpklConfigFileExists(string projectPath)
        {
            string path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

            return File.Exists(path);
        }

        public static SpklConfig GetSpklConfigFile(Project project, bool isRetry = false)
        {
            string projectPath = ProjectWorker.GetProjectPath(project);
            string path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

            try
            {
                SpklConfig spklConfig;
                using (StreamReader file = File.OpenText(path))
                {
                    JsonSerializer serializer = new JsonSerializer();
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

        public static void CreateSpklConfigFile(Project project)
        {
            TemplateHandler.AddFileFromTemplate(project, "CSharpSpklConfig\\CSharpSpklConfig", ExtensionConstants.SpklConfigFile);
        }

        public static void UpdateSpklConfigFile(string projectPath, SpklConfig spklConfig)
        {
            string text = JsonConvert.SerializeObject(spklConfig, Formatting.Indented);

            WriteSpklConfigFile(projectPath, text);
        }

        private static SpklConfig RecreateConfig(Project project, string configPath)
        {
            try
            {
                bool fileExists = FileSystem.DoesFileExist(new[] { configPath }, true);
                if (fileExists)
                    FileSystem.RenameFile(configPath);

                CreateSpklConfigFile(project);

                SpklConfig recreateConfig = GetSpklConfigFile(project, true);

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
            string path = Path.Combine(projectPath, ExtensionConstants.SpklConfigFile);

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