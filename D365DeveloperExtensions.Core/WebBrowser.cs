using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace D365DeveloperExtensions.Core
{
    public static class WebBrowser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void OpenCrmPage(CrmServiceClient client, string contentUrl)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Uri crmUri = GetBaseCrmUrlFomClient(client);

            Uri url = new Uri(crmUri, contentUrl);

            OpenPage(url.ToString());
        }

        public static Uri GetBaseCrmUrlFomClient(CrmServiceClient client)
        {
            IEnumerable<KeyValuePair<EndpointType, string>> endpoint =
                client.ConnectedOrgPublishedEndpoints.Where(k => k.Key == EndpointType.WebApplication);

            Uri crmUri = new Uri(endpoint.First().Value);

            return crmUri;
        }

        public static void OpenUrl(string contentUrl)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            OpenPage(contentUrl);
        }

        private static void OpenPage(string contentUrl)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                bool useInternalBrowser = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseInternalBrowser);

                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                if (useInternalBrowser) //Internal VS browser
                    dte.ItemOperations.Navigate(contentUrl);
                else                   //User's default browser
                    System.Diagnostics.Process.Start(contentUrl);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }
    }
}