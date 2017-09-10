using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public static class ModelBuilder
    {
        public static ObservableCollection<WebResourceItem> CreateWebResourceItemView(EntityCollection webResources, string projectName)
        {
            ObservableCollection<WebResourceItem> webResourceItems = new ObservableCollection<WebResourceItem>();

            foreach (Entity webResource in webResources.Entities)
            {
                WebResourceItem webResourceItem = new WebResourceItem
                {
                    WebResourceId = (Guid)webResource.GetAttributeValue<AliasedValue>("webresource.webresourceid").Value,
                    Name = webResource.GetAttributeValue<AliasedValue>("webresource.name").Value.ToString(),
                    IsManaged = (bool)webResource.GetAttributeValue<AliasedValue>("webresource.ismanaged").Value,
                    TypeName = Crm.WebResource.GetWebResourceTypeNameByNumber(((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value.ToString()),
                    Type = ((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value,
                    SolutionId = webResource.GetAttributeValue<EntityReference>("solutionid").Id
                };

                bool hasDisplayName = webResource.Attributes.TryGetValue("webresource.displayname", out var displayName);
                if (hasDisplayName)
                    webResourceItem.DisplayName = ((AliasedValue)displayName).Value.ToString();

                webResourceItems.Add(webResourceItem);
            }

            return webResourceItems;
        }

        public static List<CrmSolution> CreateCrmSolutionView(EntityCollection solutions)
        {
            List<CrmSolution> crmSolutions = new List<CrmSolution>();

            foreach (Entity entity in solutions.Entities)
            {
                CrmSolution solution = new CrmSolution
                {
                    SolutionId = entity.Id,
                    Name = entity.GetAttributeValue<string>("friendlyname"),
                    Prefix = entity.GetAttributeValue<AliasedValue>("publisher.customizationprefix").Value.ToString(),
                    UniqueName = entity.GetAttributeValue<string>("uniquename")
                };

                crmSolutions.Add(solution);
            }

            crmSolutions = SortSolutions(crmSolutions);

            return crmSolutions;
        }

        private static List<CrmSolution> SortSolutions(List<CrmSolution> solutions)
        {
            //Default on top
            var i = solutions.FindIndex(s => s.SolutionId == ExtensionConstants.DefaultSolutionId);
            var item = solutions[i];
            solutions.RemoveAt(i);
            solutions.Insert(0, item);

            return solutions;
        }

        public static WebResourceItem WebResourceItemFromNew(NewWebResource newWebResource, Guid solutionId)
        {
            WebResourceItem webResourceItem = new WebResourceItem
            {
                WebResourceId = newWebResource.NewId,
                Name = newWebResource.NewName,
                DisplayName = newWebResource.NewDisplayName,
                TypeName = Crm.WebResource.GetWebResourceTypeNameByNumber(newWebResource.NewType.ToString()),
                Type = newWebResource.NewType,
                SolutionId = solutionId
            };

            return webResourceItem;
        }
    }
}
