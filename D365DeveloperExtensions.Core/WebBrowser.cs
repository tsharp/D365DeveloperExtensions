using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;

namespace D365DeveloperExtensions.Core
{
    public static class WebBrowser
    {
        public static void OpenCrmPage(DTE dte, CrmServiceClient client, string contentUrl)
        {
            Uri crmUri = GetBaseCrmUrlFomClient(client);

            Uri url = new Uri(crmUri, contentUrl);

            OpenPage(dte, url.ToString());
        }

        public static Uri GetBaseCrmUrlFomClient(CrmServiceClient client)
        {
            IEnumerable<KeyValuePair<EndpointType, string>> endpoint =
                client.ConnectedOrgPublishedEndpoints.Where(k => k.Key == EndpointType.WebApplication);

            Uri crmUri = new Uri(endpoint.First().Value);

            return crmUri;
        }

        public static void OpenUrl(DTE dte, string contentUrl)
        {
            OpenPage(dte, contentUrl);
        }

        private static void OpenPage(DTE dte, string contentUrl)
        {
            bool useInternalBrowser = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseInternalBrowser);

            if (useInternalBrowser) //Internal VS browser
                dte.ItemOperations.Navigate(contentUrl);
            else                   //User's default browser
                System.Diagnostics.Process.Start(contentUrl);
        }
    }
}