using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Config
{
    public static class Mapping
    {
        public static ObservableCollection<WebResourceItem> HandleSpklMappings(Project project, string profile, ObservableCollection<WebResourceItem> webResourceItems)
        {
            SpklConfig spklConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetSpklConfigFile(project);

            List<SpklConfigWebresourceFile> mappingsToRemove = new List<SpklConfigWebresourceFile>();

            var mappedWebResources = GetSpklConfigWebresourceFiles(profile, spklConfig);
            if (mappedWebResources == null)
                return webResourceItems;

            //Remove mappings where CRM web resource was deleted
            foreach (SpklConfigWebresourceFile spklConfigWebresourceFile in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.Name == spklConfigWebresourceFile.uniquename) != 0)
                    continue;

                mappingsToRemove.Add(spklConfigWebresourceFile);
            }

            mappedWebResources = mappedWebResources.Except(mappingsToRemove).ToList();

            //Add bound file & description from mapping
            foreach (SpklConfigWebresourceFile spklConfigWebresourceFile in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.Name == spklConfigWebresourceFile.uniquename) <= 0)
                    continue;

                List<WebResourceItem> matches = webResourceItems
                    .Where(w => w.Name == spklConfigWebresourceFile.uniquename).ToList();
                foreach (WebResourceItem match in matches)
                {
                    match.Description = string.IsNullOrEmpty(spklConfigWebresourceFile.description)
                        ? null
                        : spklConfigWebresourceFile.description;
                    match.BoundFile = spklConfigWebresourceFile.file;
                }
            }

            //Remove mappings where project file was deleted
            foreach (SpklConfigWebresourceFile spklConfigWebresourceFile in mappedWebResources)
            {
                string mappedFilePath = FileSystem.BoundFileToLocalPath(spklConfigWebresourceFile.file, project.FullName);
                if (File.Exists(mappedFilePath))
                    continue;

                mappingsToRemove.Add(spklConfigWebresourceFile);
                webResourceItems.Where(w => w.BoundFile == spklConfigWebresourceFile.file).ToList().ForEach(w => w.BoundFile = null);
            }

            if (mappingsToRemove.Count > 0)
                RemoveSpklMappings(spklConfig, project, profile, mappingsToRemove);

            return webResourceItems;
        }

        public static void AddOrUpdateSpklMapping(Project project, string profile, WebResourceItem webResourceItem)
        {
            SpklConfig spklConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetSpklConfigFile(project);

            List<SpklConfigWebresourceFile> spklConfigWebresourceFiles = GetSpklConfigWebresourceFiles(profile, spklConfig) ??
                                                                         new List<SpklConfigWebresourceFile>();

            SpklConfigWebresourceFile spklConfigWebresourceFile =
                spklConfigWebresourceFiles.FirstOrDefault(w => w.uniquename == webResourceItem.Name);

            if (spklConfigWebresourceFile == null)
                AddSpklMapping(spklConfig, project, profile, webResourceItem);
            else
                UpdateSpklMapping(spklConfig, project, profile, webResourceItem);
        }

        private static void UpdateSpklMapping(SpklConfig spklConfig, Project project, string profile, WebResourceItem webResourceItem)
        {
            List<SpklConfigWebresourceFile> spklConfigWebresourceFiles = GetSpklConfigWebresourceFiles(profile, spklConfig);
            if (spklConfigWebresourceFiles == null)
                return;

            if (!string.IsNullOrEmpty(webResourceItem.BoundFile))
            {
                foreach (SpklConfigWebresourceFile spklConfigWebresourceFile in
                    spklConfigWebresourceFiles.Where(w => w.uniquename == webResourceItem.Name))
                {
                    spklConfigWebresourceFile.file = webResourceItem.BoundFile;
                    spklConfigWebresourceFile.description = webResourceItem.Description;
                }
            }
            else
                spklConfigWebresourceFiles.RemoveAll(w => w.uniquename == webResourceItem.Name);

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.webresources[0].files = spklConfigWebresourceFiles;
            else
            {
                WebresourceDeployConfig webresourceDeployConfig = spklConfig.webresources.FirstOrDefault(w => w.profile == profile);
                if (webresourceDeployConfig != null)
                    webresourceDeployConfig.files = spklConfigWebresourceFiles;
            }

            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        public static void RemoveSpklMappings(SpklConfig spklConfig, Project project, string profile, List<SpklConfigWebresourceFile> crmDexExConfigWebResources)
        {
            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.webresources[0].files
                    .RemoveAll(w => crmDexExConfigWebResources.Any(m => m.uniquename == w.uniquename));
            else
                spklConfig.webresources.FirstOrDefault(w => w.profile == profile)?.files
                    .RemoveAll(w => crmDexExConfigWebResources.Any(m => m.uniquename == w.uniquename));

            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        private static void AddSpklMapping(SpklConfig spklConfig, Project project, string profile, WebResourceItem webResourceItem)
        {
            SpklConfigWebresourceFile spklConfigWebresourceFile = new SpklConfigWebresourceFile
            {
                uniquename = webResourceItem.Name,
                file = webResourceItem.BoundFile,
                description = webResourceItem.Description
            };

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.webresources[0].files.Add(spklConfigWebresourceFile);
            else
                spklConfig.webresources.FirstOrDefault(w => w.profile == profile)?.files.Add(spklConfigWebresourceFile);

            string projectPath = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectPath(project);
            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        private static List<SpklConfigWebresourceFile> GetSpklConfigWebresourceFiles(string profile, SpklConfig spklConfig)
        {
            List<SpklConfigWebresourceFile> spklConfigWebresourceFiles = profile.StartsWith(ExtensionConstants.NoProfilesText)
                ? spklConfig.webresources[0].files
                : spklConfig.webresources.FirstOrDefault(w => w.profile == profile)?.files;

            return spklConfigWebresourceFiles;
        }
    }
}