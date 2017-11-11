using EnvDTE;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CrmDeveloperExtensions2.Core
{
    public static class WebBrowser
    {
        public static void OpenCrmPage(DTE dte, CrmServiceClient client, string contentUrl)
        {
            bool useInternalBrowser = UserOptionsGrid.GetUseInternalBrowser(dte);

            Uri crmUri = GetBaseCrmUrlFomClient(client);

            Uri url = new Uri(crmUri, contentUrl);

            if (useInternalBrowser) //Internal VS browser
                dte.ItemOperations.Navigate(url.ToString());
            else //User's default browser
                System.Diagnostics.Process.Start(url.ToString());
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
            bool useInternalBrowser = UserOptionsGrid.GetUseInternalBrowser(dte);

            if (useInternalBrowser) //Internal VS browser
                dte.ItemOperations.Navigate(contentUrl);
            else //User's default browser
                System.Diagnostics.Process.Start(contentUrl);
        }
    }
}