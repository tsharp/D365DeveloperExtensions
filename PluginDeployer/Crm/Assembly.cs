using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using PluginDeployer.ViewModels;
using System;
using System.ServiceModel;

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
                            EntityAlias = "plugintype"
                        },
                        new LinkEntity
                        {
                            Columns = new ColumnSet("solutionid"),
                            LinkFromEntityName = "pluginassembly",
                            LinkFromAttributeName = "pluginassemblyid",
                            LinkToEntityName = "solutioncomponent",
                            LinkToAttributeName = "objectid",
                            EntityAlias = "solutioncomponent",
                        }
                    },
                    Orders =
                    {
                        new OrderExpression
                        {
                            AttributeName = "name",
                            OrderType = OrderType.Ascending
                        }
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

        public static bool UpdateCrmAssembly(CrmServiceClient client, CrmAssembly crmAssembly, string assebmlyPath)
        {
            try
            {
                Entity assembly = new Entity("pluginassembly")
                {
                    Id = crmAssembly.AssemblyId,
                    ["content"] = Convert.ToBase64String(CrmDeveloperExtensions2.Core.FileSystem.GetFileBytes(assebmlyPath))
                };

                client.Update(assembly);

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Updating Assembly In CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Updating Assembly In CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }
    }
}
