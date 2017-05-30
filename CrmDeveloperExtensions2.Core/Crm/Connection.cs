using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace CrmDeveloperExtensions2.Core.Crm
{
    public static class Connection
    {
        public static string RetrieveOrganizationId(CrmServiceClient service)
        {
            WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());
            return response.OrganizationId.ToString();
        }
    }
}
