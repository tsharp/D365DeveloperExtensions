using CrmDeveloperExtensions2.Core.Resources;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;

namespace CrmDeveloperExtensions2.Core.Crm
{
    public class Publish
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool PublishAllCustomizations(CrmServiceClient client)
        {
            try
            {
                PublishAllXmlRequest request = new PublishAllXmlRequest();

                client.Execute(request);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorPublishing, ex);

                return false;
            }
        }
    }
}