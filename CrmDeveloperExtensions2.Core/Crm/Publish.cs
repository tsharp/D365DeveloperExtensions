using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.ServiceModel;

namespace CrmDeveloperExtensions2.Core.Crm
{
    public class Publish
    {
        public static bool PublishAllCustomizations(CrmServiceClient client)
        {
            try
            {
                PublishAllXmlRequest request = new PublishAllXmlRequest();

                client.Execute(request);

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow("Error Publishing Customizations To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Publishing Customizations To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }
    }
}
