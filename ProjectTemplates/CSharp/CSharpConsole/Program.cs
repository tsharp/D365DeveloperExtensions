$if$ ($useXrmToolingClientUsing$ == 1)using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace CRMSimpleConsole
{
    class Program
    {
        private static CrmServiceClient _client;

        public static void Main(string[] args)
        {
            try
            {
                using (_client = new CrmServiceClient(ConfigurationManager.ConnectionStrings["CRMConnectionString"].ConnectionString))
                {
                    //Do stuff
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                string message = ex.Message;
                throw;
            }
        }
    }
}$else$using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;

namespace CRMSimpleConsole
{
    class Program
    {
        private static OrganizationService _orgService;
        static void Main(string[] args)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(
                    ConfigurationManager.ConnectionStrings["CRMConnectionString"].ConnectionString);

                using (_orgService = new OrganizationService(connection))
                {
                    //Do stuff
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                string message = ex.Message;
                throw;
            }
        }
    }
}$endif$