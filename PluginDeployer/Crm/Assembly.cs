using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace PluginDeployer.Crm
{
    public static class Assembly
    {
        public static EntityCollection RetrieveAssembliesFromCrm(CrmServiceClient client)
        {
            try
            {
                QueryExpression query = new QueryExpression
                {
                    EntityName = "pluginassembly",
                    ColumnSet = new ColumnSet("pluginassemblyid", "name", "version"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "ismanaged",
                                Operator = ConditionOperator.Equal,
                                Values = { false }
                            }
                        }
                    },
                    LinkEntities =
                    {
                        new LinkEntity
                        {
                            Columns = new ColumnSet("isworkflowactivity"),
                            LinkFromEntityName = "pluginassembly",
                            LinkFromAttributeName = "pluginassemblyid",
                            LinkToEntityName = "plugintype",
                            LinkToAttributeName = "pluginassemblyid",
                            EntityAlias = "plugintype",
                            JoinOperator = JoinOperator.LeftOuter
                        }
                        //,
                        //new LinkEntity
                        //{
                        //    Columns = new ColumnSet("solutionid"),
                        //    LinkFromEntityName = "pluginassembly",
                        //    LinkFromAttributeName = "pluginassemblyid",
                        //    LinkToEntityName = "solutioncomponent",
                        //    LinkToAttributeName = "solutioncomponentid",
                        //    EntityAlias = "solutioncomponent",
                        //}
                    }
                };

                EntityCollection assemblies = client.RetrieveMultiple(query);

                return assemblies;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Assemblies From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Assemblies From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }
    }
}
