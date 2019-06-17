using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public static class ModelBuilder
    {
        public static ObservableCollection<WebResourceItem> CreateWebResourceItemView(EntityCollection webResources, string projectName)
        {
            var webResourceItems = new ObservableCollection<WebResourceItem>();

            foreach (var webResource in webResources.Entities)
            {
                var webResourceItem = new WebResourceItem
                {
                    WebResourceId = (Guid)webResource.GetAttributeValue<AliasedValue>("webresource.webresourceid").Value,
                    Name = webResource.GetAttributeValue<AliasedValue>("webresource.name").Value.ToString(),
                    IsManaged = (bool)webResource.GetAttributeValue<AliasedValue>("webresource.ismanaged").Value,
                    TypeName = WebResourceTypes.GetWebResourceTypeNameByNumber(((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value.ToString()),
                    Type = ((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value,
                    SolutionId = webResource.GetAttributeValue<EntityReference>("solutionid").Id
                };

                var hasDescription = webResource.Attributes.TryGetValue("webresource.description", out var description);
                if (hasDescription)
                {
                    webResourceItem.Description = ((AliasedValue)description).Value.ToString();
                    webResourceItem.PreviousDescription = webResourceItem.Description;
                }

                var hasDisplayName = webResource.Attributes.TryGetValue("webresource.displayname", out var displayName);
                if (hasDisplayName)
                    webResourceItem.DisplayName = ((AliasedValue)displayName).Value.ToString();

                webResourceItems.Add(webResourceItem);
            }

            webResourceItems = new ObservableCollection<WebResourceItem>(webResourceItems.OrderBy(w => w.Name));

            return webResourceItems;
        }

        public static List<CrmSolution> CreateCrmSolutionView(EntityCollection solutions)
        {
            var crmSolutions = new List<CrmSolution>();

            foreach (var entity in solutions.Entities)
            {
                var solution = new CrmSolution
                {
                    SolutionId = entity.Id,
                    Name = entity.GetAttributeValue<string>("friendlyname"),
                    Prefix = entity.GetAttributeValue<AliasedValue>("publisher.customizationprefix").Value.ToString(),
                    UniqueName = entity.GetAttributeValue<string>("uniquename"),
                    NameVersion = $"{entity.GetAttributeValue<string>("friendlyname")} {entity.GetAttributeValue<string>("version")}"
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
            var webResourceItem = new WebResourceItem
            {
                WebResourceId = newWebResource.NewId,
                Name = newWebResource.NewName,
                DisplayName = newWebResource.NewDisplayName,
                TypeName = WebResourceTypes.GetWebResourceTypeNameByNumber(newWebResource.NewType.ToString()),
                Type = newWebResource.NewType,
                SolutionId = solutionId,
                Description = newWebResource.NewDescription,
                PreviousDescription = newWebResource.NewDescription,
                IsManaged = false
            };

            return webResourceItem;
        }
    }
}