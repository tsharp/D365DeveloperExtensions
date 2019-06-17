using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Config;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Vs;
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
            var spklConfig = D365DeveloperExtensions.Core.Config.Mapping.GetSpklConfigFile(project);

            var mappedWebResources = GetSpklConfigWebresourceFiles(profile, spklConfig);
            if (mappedWebResources == null)
                return webResourceItems;

            //Lock mappings where CRM web resource was deleted
            foreach (var spklConfigWebresourceFile in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.Name == spklConfigWebresourceFile.uniquename) == 0)
                {
                    webResourceItems.Where(w => w.Name == spklConfigWebresourceFile.uniquename).ToList()
                        .ForEach(ww => ww.Locked = true);
                }
            }

            //Add bound file & description from mapping
            foreach (var spklConfigWebresourceFile in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.Name == spklConfigWebresourceFile.uniquename) <= 0)
                    continue;

                var matches = webResourceItems.Where(w => w.Name == spklConfigWebresourceFile.uniquename).ToList();
                foreach (var match in matches)
                {
                    match.Description = string.IsNullOrEmpty(spklConfigWebresourceFile.description)
                        ? null
                        : spklConfigWebresourceFile.description;
                    match.BoundFile = string.IsNullOrEmpty(spklConfigWebresourceFile.ts)
                        ? spklConfigWebresourceFile.file
                        : spklConfigWebresourceFile.ts;
                }
            }

            //Lock mappings where project file was deleted
            foreach (var spklConfigWebresourceFile in mappedWebResources)
            {
                var mappedFilePath = FileSystem.BoundFileToLocalPath(spklConfigWebresourceFile.file, project.FullName);
                var relativePath = ProjectItemWorker.GetRelativePathFromPath(ProjectWorker.GetProjectPath(project), mappedFilePath);
                var inProject = ProjectWorker.IsFileInProjectFile(project.FullName, relativePath);
                if (File.Exists(mappedFilePath) && inProject)
                    continue;

                if (!string.IsNullOrEmpty(spklConfigWebresourceFile.ts))
                    mappedFilePath = FileSystem.BoundFileToLocalPath(spklConfigWebresourceFile.ts, project.FullName);
                if (File.Exists(mappedFilePath) && inProject)
                    continue;

                webResourceItems.Where(w => w.BoundFile == spklConfigWebresourceFile.file).ToList().ForEach(w => w.Locked = true);
            }

            return webResourceItems;
        }

        public static void AddOrUpdateSpklMapping(Project project, string profile, WebResourceItem webResourceItem)
        {
            var spklConfig = D365DeveloperExtensions.Core.Config.Mapping.GetSpklConfigFile(project);

            var spklConfigWebresourceFiles = GetSpklConfigWebresourceFiles(profile, spklConfig) ??
                                                                         new List<SpklConfigWebresourceFile>();

            var spklConfigWebresourceFile =
                spklConfigWebresourceFiles.FirstOrDefault(w => w.uniquename == webResourceItem.Name);

            if (spklConfigWebresourceFile == null)
                AddSpklMapping(spklConfig, project, profile, webResourceItem);
            else
                UpdateSpklMapping(spklConfig, project, profile, webResourceItem);
        }

        private static void UpdateSpklMapping(SpklConfig spklConfig, Project project, string profile, WebResourceItem webResourceItem)
        {
            var spklConfigWebresourceFiles = GetSpklConfigWebresourceFiles(profile, spklConfig);
            if (spklConfigWebresourceFiles == null)
                return;

            if (!string.IsNullOrEmpty(webResourceItem.BoundFile))
            {
                var isTs = WebResourceTypes.GetExtensionType(webResourceItem.BoundFile) == D365DeveloperExtensions.Core.Enums.FileExtensionType.Ts;
                foreach (var spklConfigWebresourceFile in
                    spklConfigWebresourceFiles.Where(w => w.uniquename == webResourceItem.Name))
                {
                    spklConfigWebresourceFile.file = webResourceItem.BoundFile;
                    spklConfigWebresourceFile.description = webResourceItem.Description;
                    spklConfigWebresourceFile.ts = null;

                    if (!isTs)
                        continue;

                    spklConfigWebresourceFile.file = TsHelper.GetJsForTsPath(webResourceItem.BoundFile, project);
                    spklConfigWebresourceFile.ts = webResourceItem.BoundFile;
                }
            }
            else
                spklConfigWebresourceFiles.RemoveAll(w => w.uniquename == webResourceItem.Name);

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.webresources[0].files = spklConfigWebresourceFiles;
            else
            {
                var webresourceDeployConfig = spklConfig.webresources.FirstOrDefault(w => w.profile == profile);
                if (webresourceDeployConfig != null)
                    webresourceDeployConfig.files = spklConfigWebresourceFiles;
            }

            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        public static void RemoveSpklMappings(SpklConfig spklConfig, Project project, string profile, List<SpklConfigWebresourceFile> crmDexExConfigWebResources)
        {
            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);

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
            var spklConfigWebresourceFile = new SpklConfigWebresourceFile
            {
                uniquename = webResourceItem.Name,
                file = webResourceItem.BoundFile,
                description = webResourceItem.Description
            };

            if (WebResourceTypes.GetExtensionType(webResourceItem.BoundFile) == D365DeveloperExtensions.Core.Enums.FileExtensionType.Ts)
            {
                spklConfigWebresourceFile.file = TsHelper.GetJsForTsPath(webResourceItem.BoundFile, project);
                spklConfigWebresourceFile.ts = webResourceItem.BoundFile;
            }

            if (profile.StartsWith(ExtensionConstants.NoProfilesText))
                spklConfig.webresources[0].files.Add(spklConfigWebresourceFile);
            else
                spklConfig.webresources.FirstOrDefault(w => w.profile == profile)?.files.Add(spklConfigWebresourceFile);

            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            ConfigFile.UpdateSpklConfigFile(projectPath, spklConfig);
        }

        private static List<SpklConfigWebresourceFile> GetSpklConfigWebresourceFiles(string profile, SpklConfig spklConfig)
        {
            var spklConfigWebresourceFiles = profile.StartsWith(ExtensionConstants.NoProfilesText)
                ? spklConfig.webresources[0].files
                : spklConfig.webresources.FirstOrDefault(w => w.profile == profile)?.files;

            return spklConfigWebresourceFiles;
        }
    }
}