using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Reflection;
using System.Web.UI.WebControls;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class ConfigFile
    {
        private static readonly string ConfigFileName = ExtensionConstants.SpklConfigFile;

        public static bool SpklConfigFileExists(string projectPath)
        {
            string path = Path.Combine(projectPath, ConfigFileName);

            return File.Exists(path);
        }

        public static SpklConfig GetSpklConfigFile(string projectPath)
        {
            string path = Path.Combine(projectPath, ConfigFileName);
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
            catch
            {
                throw new Exception("Unable to read or deserialize config file");
            }
        }

        public static void CreateSpklConfigFile(Project project)
        {
            TemplateHandler.AddFileFromTemplate(project, "SpklConfig\\SpklConfig", ConfigFileName);
        }

        public static void UpdateSpklConfigFile(string projectPath, SpklConfig spklConfig)
        {
            string text = JsonConvert.SerializeObject(spklConfig, Formatting.Indented);

            WriteSpklConfigFile(projectPath, text);
        }

        private static void WriteSpklConfigFile(string projectPath, string text)
        {
            try
            {
                string path = Path.Combine(projectPath, ConfigFileName);
                File.WriteAllText(path, text);
            }
            catch
            {
                MessageBox.Show("Error writing config file");
            }
        }
    }
}