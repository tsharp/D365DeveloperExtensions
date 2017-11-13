using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Config
{
    public static class ConfigFile
    {
        private static readonly string ConfigFileName = ExtensionConstants.SpklConfigFile;

        public static bool ConfigFileExists(string solutionPath)
        {
            DirectoryInfo directory = FileSystem.GetDirectory(solutionPath);

            string path = $"{directory.FullName}\\{ConfigFileName}";

            return File.Exists(path);
        }

        public static bool SpklConfigFileExists(string projectPath)
        {
            string path = Path.Combine(projectPath, ConfigFileName);

            return File.Exists(path);
        }

        public static CrmDexExConfig GetConfigFile(string solutionPath)
        {
            DirectoryInfo directory = FileSystem.GetDirectory(solutionPath);

            try
            {
                CrmDexExConfig config;
                using (StreamReader file = File.OpenText($"{directory.FullName}\\{ConfigFileName}"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    config = (CrmDexExConfig)serializer.Deserialize(file, typeof(CrmDexExConfig));
                }

                return config;
            }
            catch
            {
                throw new Exception("Unable to read or deserialize config file");
            }
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

        public static CrmDexExConfig CreateConfigFile(Guid organizationId, string projectUniqueName, string solutionPath)
        {
            CrmDexExConfig crmDexExConfig = new CrmDexExConfig
            {
                CrmDevExConfigOrgMaps = new List<CrmDevExConfigOrgMap>
                {
                    new CrmDevExConfigOrgMap
                    {
                        OrganizationId = organizationId,
                        ProjectUniqueName = projectUniqueName
                    }
                }
            };

            string text = JsonConvert.SerializeObject(crmDexExConfig, Formatting.Indented);

            WriteConfigFile(solutionPath, text);

            return crmDexExConfig;
        }

        public static SpklConfig CreateSpklConfigFile(Project project)
        {
            SpklConfig spklConfig = null;
            try
            {
                string codebase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(codebase, UriKind.Absolute);
                string path = Path.GetDirectoryName(uri.LocalPath);

                if (string.IsNullOrEmpty(path))
                {
                    OutputLogger.WriteToOutputWindow("Error finding extension template directory", MessageType.Error);
                    return null;
                }

                var templatePath = Path.Combine(path, @"ItemTemplates\CSharp\Crm DevEx\1033\SpklConfig\SpklConfig.vstemplate");

                project.ProjectItems.AddFromTemplate(templatePath, ConfigFileName);

                spklConfig = GetSpklConfigFile(Vs.ProjectWorker.GetProjectPath(project));
            }
            catch
            {
                MessageBox.Show("Error creating config file");
            }

            return spklConfig;
        }

        public static void UpdateSpklConfigFile(string projectPath, SpklConfig spklConfig)
        {
            string text = JsonConvert.SerializeObject(spklConfig, Formatting.Indented);

            WriteSpklConfigFile(projectPath, text);
        }

        public static void UpdateConfigFile(string solutionPath, CrmDexExConfig crmDexExConfig)
        {
            string text = JsonConvert.SerializeObject(crmDexExConfig, Formatting.Indented);

            WriteConfigFile(solutionPath, text);
        }

        private static void WriteConfigFile(string solutionPath, string text)
        {
            DirectoryInfo directory = FileSystem.GetDirectory(solutionPath);

            try
            {
                File.WriteAllText($"{directory.FullName}\\{ConfigFileName}", text);
            }
            catch
            {
                MessageBox.Show("Error writing config file");
            }
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