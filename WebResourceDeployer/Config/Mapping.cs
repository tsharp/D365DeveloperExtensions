using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Config
{
    public static class Mapping
    {
        public static ObservableCollection<WebResourceItem> HandleMappings(string solutionPath, Project project, ObservableCollection<WebResourceItem> webResourceItems, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);
            if (crmDevExConfigOrgMap.WebResources == null)
                crmDevExConfigOrgMap.WebResources = new List<CrmDexExConfigWebResource>();

            List<CrmDexExConfigWebResource> mappingsToRemove = new List<CrmDexExConfigWebResource>();

            //Remove mappings where CRM web resource was deleted
            List<CrmDexExConfigWebResource> mappedWebResources = crmDevExConfigOrgMap.WebResources;

            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId) != 0)
                    continue;

                mappingsToRemove.Add(crmDexExConfigWebResource);
                mappedWebResources.Remove(crmDexExConfigWebResource);
            }

            //Add bound file from mapping
            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
            {
                if (webResourceItems.Count(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId) > 0)
                    webResourceItems.Where(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId).ToList()
                        .ForEach(w => w.BoundFile = crmDexExConfigWebResource.File);
            }

            //Remove mappings where project file was deleted
            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
            {
                string mappedFilePath = FileSystem.BoundFileToLocalPath(crmDexExConfigWebResource.File, project.FullName);
                if (!File.Exists(mappedFilePath))
                {
                    mappingsToRemove.Add(crmDexExConfigWebResource);
                    webResourceItems.Where(w => w.BoundFile == crmDexExConfigWebResource.File).ToList().ForEach(w => w.BoundFile = null);
                }
            }

            if (mappingsToRemove.Count > 0)
                RemoveMappings(solutionPath, project, mappingsToRemove, organizationId);

            return webResourceItems;
        }

        public static void AddOrUpdateMapping(string solutionPath, Project project, WebResourceItem webResourceItem, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            List<CrmDexExConfigWebResource> crmDexExConfigWebResources = crmDevExConfigOrgMap.WebResources;
            if (crmDexExConfigWebResources == null)
                crmDevExConfigOrgMap.WebResources = new List<CrmDexExConfigWebResource>();

            CrmDexExConfigWebResource crmDexExConfigWebResource =
                crmDevExConfigOrgMap.WebResources.FirstOrDefault(w => w.WebResourceId == webResourceItem.WebResourceId);

            if (crmDexExConfigWebResource == null)
                AddMapping(crmDexExConfig, solutionPath, project, webResourceItem, organizationId);
            else
                UpdateMapping(crmDexExConfig, solutionPath, project, webResourceItem, organizationId);
        }

        private static void UpdateMapping(CrmDexExConfig crmDexExConfig, string solutionPath, Project project, WebResourceItem webResourceItem, Guid organizationId)
        {
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            List<CrmDexExConfigWebResource> crmDexExConfigWebResources = crmDevExConfigOrgMap.WebResources.Where(w => w.WebResourceId == webResourceItem.WebResourceId).ToList();

            crmDexExConfigWebResources.ForEach(w => w.File = webResourceItem.BoundFile);
            crmDevExConfigOrgMap.WebResources.RemoveAll(w => string.IsNullOrEmpty(w.File));

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        public static void RemoveMappings(string solutionPath, Project project, List<CrmDexExConfigWebResource> crmDexExConfigWebResources, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = CrmDeveloperExtensions2.Core.Config.Mapping.GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            crmDevExConfigOrgMap.WebResources.RemoveAll(
                w => crmDexExConfigWebResources.Any(m => m.WebResourceId == w.WebResourceId));

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        private static void AddMapping(CrmDexExConfig crmDexExConfig, string solutionPath, Project project, WebResourceItem webResourceItem, Guid organizationId)
        {
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = CrmDeveloperExtensions2.Core.Config.Mapping.GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            CrmDexExConfigWebResource crmDexExConfigWebResource = new CrmDexExConfigWebResource
            {
                WebResourceId = webResourceItem.WebResourceId,
                File = webResourceItem.BoundFile
            };

            crmDevExConfigOrgMap.WebResources.Add(crmDexExConfigWebResource);

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }
    }
}
