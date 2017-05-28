using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using EnvDTE;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using WebResourceDeployer.ViewModels;
using Menu = System.Web.UI.WebControls.Menu;

namespace WebResourceDeployer
{
    public static class Class1
    {
        public static List<WebResourceItem> CreateWebResourceItemView(EntityCollection webResources, string projectName, ObservableCollection<MenuItem> projectFolders)
        {
            List<WebResourceItem> webResourceItems = new List<WebResourceItem>();

            foreach (Entity webResource in webResources.Entities)
            {
                WebResourceItem webResourceItem = new WebResourceItem
                {
                    Publish = false,
                    WebResourceId = (Guid)webResource.GetAttributeValue<AliasedValue>("webresource.webresourceid").Value,
                    Name = webResource.GetAttributeValue<AliasedValue>("webresource.name").Value.ToString(),
                    IsManaged = (bool)webResource.GetAttributeValue<AliasedValue>("webresource.ismanaged").Value,
                    AllowPublish = false,
                    AllowCompare = false,
                    TypeName = Crm.WebResource.GetWebResourceTypeNameByNumber(((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value.ToString()),
                    Type = ((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value,
                    ProjectFolders = projectFolders,
                    SolutionId = webResource.GetAttributeValue<EntityReference>("solutionid").Id
                };

                object displayName;
                bool hasDisplayName = webResource.Attributes.TryGetValue("webresource.displayname", out displayName);
                if (hasDisplayName)
                    webResourceItem.DisplayName = webResource.GetAttributeValue<AliasedValue>("webresource.displayname").Value.ToString();

                webResourceItems.Add(webResourceItem);
            }

            return webResourceItems;
        }

        public static ObservableCollection<WebResourceItem> CreateWebResourceItemView2(EntityCollection webResources, string projectName, ObservableCollection<MenuItem> projectFolders)
        {
            ObservableCollection<WebResourceItem> webResourceItems = new ObservableCollection<WebResourceItem>();

            foreach (Entity webResource in webResources.Entities)
            {
                WebResourceItem webResourceItem = new WebResourceItem
                {
                    Publish = false,
                    WebResourceId = (Guid)webResource.GetAttributeValue<AliasedValue>("webresource.webresourceid").Value,
                    Name = webResource.GetAttributeValue<AliasedValue>("webresource.name").Value.ToString(),
                    IsManaged = (bool)webResource.GetAttributeValue<AliasedValue>("webresource.ismanaged").Value,
                    AllowPublish = false,
                    AllowCompare = false,
                    TypeName = Crm.WebResource.GetWebResourceTypeNameByNumber(((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value.ToString()),
                    Type = ((OptionSetValue)webResource.GetAttributeValue<AliasedValue>("webresource.webresourcetype").Value).Value,
                    ProjectFolders = projectFolders,
                    SolutionId = webResource.GetAttributeValue<EntityReference>("solutionid").Id
                };

                object displayName;
                bool hasDisplayName = webResource.Attributes.TryGetValue("webresource.displayname", out displayName);
                if (hasDisplayName)
                    webResourceItem.DisplayName = webResource.GetAttributeValue<AliasedValue>("webresource.displayname").Value.ToString();

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
            var i = solutions.FindIndex(s => s.SolutionId == CrmDeveloperExtensions.Core.ExtensionConstants.DefaultSolutionId);
            var item = solutions[i];
            solutions.RemoveAt(i);
            solutions.Insert(0, item);

            return solutions;
        }

        public static WebResourceItem WebResourceItemFromNew(NewWebResource newWebResource, Guid solutionId, ObservableCollection<MenuItem> projectFolders)
        {
            WebResourceItem webResourceItem = new WebResourceItem
            {
                Publish = false,
                WebResourceId = newWebResource.NewId,
                Name = newWebResource.NewName,
                DisplayName = newWebResource.NewDisplayName,
                IsManaged = false,
                AllowPublish = true,
                AllowCompare = SetAllowCompare(newWebResource.NewType),
                TypeName = Crm.WebResource.GetWebResourceTypeNameByNumber(newWebResource.NewType.ToString()),
                Type = newWebResource.NewType,
                ProjectFolders = projectFolders,
                SolutionId = solutionId
            };

            return webResourceItem;
        }

        public static bool SetAllowCompare(int type)
        {
            int[] noCompare = { 5, 6, 7, 8, 10 };
            return !noCompare.Contains(type);
        }
    }
}
