using System;
using System.ServiceModel;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace PluginTraceViewer.Crm
{
    public static class PluginTrace
    {
        public static EntityCollection RetrievePluginTracesFromCrm(CrmServiceClient client)
        {
            try
            {
                QueryExpression query = new QueryExpression
                {
                    EntityName = "plugintracelog",
                    ColumnSet = new ColumnSet("primaryentity", "correlationid", "messageblock", "messagename", "depth",
                    "performanceexecutionduration", "typename", "createdon"),
                    LinkEntities =
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "plugintracelog",
                            LinkFromAttributeName = "createdby",
                            LinkToEntityName = "systemuser",
                            LinkToAttributeName = "systemuserid",
                            EntityAlias = "createdbyuser",
                            Columns = new ColumnSet("fullname")
                        }
                    },
                    Orders =
                    {
                        new OrderExpression
                        {
                            AttributeName = "createdon",
                            OrderType = OrderType.Descending
                        }
                    }
                };

                EntityCollection traceLogs = client.RetrieveMultiple(query);

                return traceLogs;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Plug-in Trace Logs From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Plug-in Trace Logs From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }
    }
}