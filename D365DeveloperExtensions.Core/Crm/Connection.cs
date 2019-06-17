using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace D365DeveloperExtensions.Core.Crm
{
    public static class Connection
    {
        public static string RetrieveOrganizationId(CrmServiceClient service)
        {
            var response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

            return response.OrganizationId.ToString();
        }
    }
}