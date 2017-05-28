using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using CrmDeveloperExtensions.Core;
using CrmDeveloperExtensions.Core.Config;
using CrmDeveloperExtensions.Core.Models;
using EnvDTE;
using WebResourceDeployer.ViewModels;
using Newtonsoft.Json;

namespace WebResourceDeployer
{
    public static class Mapping
    {
        //public static List<WebResourceItem> HandleMappings(DTE dte, Project project, List<WebResourceItem> webResourceItems, Guid organizationId)
        public static ObservableCollection<WebResourceItem> HandleMappings(DTE dte, Project project, ObservableCollection<WebResourceItem> webResourceItems, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = GetConfigFile(dte.Solution.FullName, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);
            if (crmDevExConfigOrgMap.WebResources == null)
                crmDevExConfigOrgMap.WebResources = new List<CrmDexExConfigWebResource>();

            List<CrmDexExConfigWebResource> mappingsToRemove = new List<CrmDexExConfigWebResource>();

            //Remove mappings where CRM web resource was deleted
            List<CrmDexExConfigWebResource> mappedWebResources = crmDevExConfigOrgMap.WebResources;

            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
                if (webResourceItems.Count(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId) == 0)
                {
                    mappingsToRemove.Add(crmDexExConfigWebResource);
                    mappedWebResources.Remove(crmDexExConfigWebResource);
                }

            //Add bound file from mapping
            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
                if (webResourceItems.Count(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId) > 0)
                    webResourceItems.Where(w => w.WebResourceId == crmDexExConfigWebResource.WebResourceId).ToList().ForEach(w => w.BoundFile = crmDexExConfigWebResource.File);

            //Remove mappings where project file was deleted
            foreach (CrmDexExConfigWebResource crmDexExConfigWebResource in mappedWebResources)
            {
                string mappedFilePath = FileSystem.BoundFileToLocalPath(crmDexExConfigWebResource.File, project.FullName);
                if (!File.Exists(mappedFilePath))
                    mappingsToRemove.Add(crmDexExConfigWebResource);
            }

            if (mappingsToRemove.Count > 0)
                RemoveMappings(dte.Solution.FullName, project, mappingsToRemove, organizationId);

            return webResourceItems;
        }

        public static void AddOrUpdateMapping(string solutionPath, Project project, WebResourceItem webResourceItem, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

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
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            List<CrmDexExConfigWebResource> crmDexExConfigWebResources = crmDevExConfigOrgMap.WebResources.Where(w => w.WebResourceId == webResourceItem.WebResourceId).ToList();

            crmDexExConfigWebResources.ForEach(w => w.File = webResourceItem.BoundFile);
            crmDevExConfigOrgMap.WebResources.RemoveAll(w => string.IsNullOrEmpty(w.File));

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        public static void RemoveMappings(string solutionPath, Project project, List<CrmDexExConfigWebResource> crmDexExConfigWebResources, Guid organizationId)
        {
            CrmDexExConfig crmDexExConfig = GetConfigFile(solutionPath, project.UniqueName, organizationId);
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            crmDevExConfigOrgMap.WebResources.RemoveAll(
                w => crmDexExConfigWebResources.Any(m => m.WebResourceId == w.WebResourceId));

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        private static void AddMapping(CrmDexExConfig crmDexExConfig, string solutionPath, Project project, WebResourceItem webResourceItem, Guid organizationId)
        {
            CrmDevExConfigOrgMap crmDevExConfigOrgMap = GetOrgMap(ref crmDexExConfig, organizationId, project.UniqueName);

            CrmDexExConfigWebResource crmDexExConfigWebResource = new CrmDexExConfigWebResource
            {
                WebResourceId = webResourceItem.WebResourceId,
                File = webResourceItem.BoundFile
            };

            crmDevExConfigOrgMap.WebResources.Add(crmDexExConfigWebResource);

            ConfigFile.UpdateConfigFile(solutionPath, crmDexExConfig);
        }

        private static CrmDexExConfig GetConfigFile(string solutionPath, string projectUniqueName, Guid organizationId)
        {
            return !ConfigFile.ConfigFileExists(solutionPath) ?
                ConfigFile.CreateConfigFile(organizationId, projectUniqueName, solutionPath) :
                ConfigFile.GetConfigFile(solutionPath);
        }

        private static CrmDevExConfigOrgMap GetOrgMap(ref CrmDexExConfig crmDexExConfig, Guid organizationId, string projectUniqueName)
        {
            CrmDevExConfigOrgMap orgMap = crmDexExConfig.CrmDevExConfigOrgMaps.FirstOrDefault(o => o.OrganizationId == organizationId);
            if (orgMap != null)
                return orgMap;

            orgMap = new CrmDevExConfigOrgMap
            {
                OrganizationId = organizationId,
                ProjectUniqueName = projectUniqueName,
                WebResources = new List<CrmDexExConfigWebResource>()
            };

            crmDexExConfig.CrmDevExConfigOrgMaps.Add(orgMap);

            return orgMap;
        }
    }
}
