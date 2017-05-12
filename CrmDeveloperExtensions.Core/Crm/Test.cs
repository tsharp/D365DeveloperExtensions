using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace CrmDeveloperExtensions.Core.Crm
{
    public static class Test
    {
        public static string DoWhoAmI(CrmServiceClient service)
        {
            WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());
            return response.UserId.ToString();
        }
    }
}
