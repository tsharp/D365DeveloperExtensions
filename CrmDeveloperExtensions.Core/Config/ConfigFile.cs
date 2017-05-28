using CrmDeveloperExtensions.Core.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace CrmDeveloperExtensions.Core.Config
{
    public static class ConfigFile
    {
        private static readonly string ConfigFileName = Resources.Resource.ConfigFileName;

        public static bool ConfigFileExists(string solutionPath)
        {
            DirectoryInfo directory = FileSystem.GetDirectory(solutionPath);

            string path = $"{directory.FullName}\\{ConfigFileName}";

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
    }
}
