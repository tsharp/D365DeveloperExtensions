using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using PluginDeployer.Spkl;
using PluginDeployer.ViewModels;
using System;
using System.Linq;
using System.ServiceModel;

namespace PluginDeployer.Crm
{
    public class Assembly
    {
        public static Entity RetrieveAssemblyFromCrm(CrmServiceClient client, string assemblyName)
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
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = { assemblyName }
                            }
                        }
                    }
                };

                EntityCollection assemblies = client.RetrieveMultiple(query);

                return assemblies.Entities.FirstOrDefault();
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Assembly From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Assembly From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }

        public static Guid UpdateCrmAssembly(CrmServiceClient client, CrmAssembly crmAssembly)
        {
            try
            {
                Entity assembly = new Entity("pluginassembly")
                {
                    ["content"] = Convert.ToBase64String(CrmDeveloperExtensions2.Core.FileSystem.GetFileBytes(crmAssembly.AssemblyPath)),
                    ["name"] = crmAssembly.Name,
                    ["culture"] = crmAssembly.Culture,
                    ["version"] = crmAssembly.Version,
                    ["publickeytoken"] = crmAssembly.PublicKeyToken,
                    ["sourcetype"] = new OptionSetValue(0), // database
                    ["isolationmode"] = crmAssembly.IsolationMode == IsolationModeEnum.Sandbox ? new OptionSetValue(2) : new OptionSetValue(1) // 1= none, 2 = sandbox
                };

                if (crmAssembly.AssemblyId == Guid.Empty)
                    return client.Create(assembly);

                assembly.Id = crmAssembly.AssemblyId;
                client.Update(assembly);
                return crmAssembly.AssemblyId;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                //TODO: change message based on create or update
                OutputLogger.WriteToOutputWindow(
                    "Error Updating Assembly In CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return Guid.Empty;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Updating Assembly In CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return Guid.Empty;
            }
        }

        public static bool AddAssemblyToSolution(CrmServiceClient client, Guid assemblyId, string uniqueName)
        {
            try
            {
                AddSolutionComponentRequest scRequest = new AddSolutionComponentRequest
                {
                    ComponentType = 91,
                    SolutionUniqueName = uniqueName,
                    ComponentId = assemblyId
                };
                AddSolutionComponentResponse response =
                    (AddSolutionComponentResponse)client.Execute(scRequest);

                OutputLogger.WriteToOutputWindow("New Assembly Added To Solution: " + response.id, MessageType.Info);

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                //TODO: change message based on create or update
                OutputLogger.WriteToOutputWindow(
                    "Error adding assembly to solution: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error adding assembly to solution: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }

        public static bool IsAssemblyInSolution(CrmServiceClient client, string assemblyName, string uniqueName)
        {
            try
            {
                FetchExpression query = new FetchExpression($@"<fetch>
                                                          <entity name='solutioncomponent'>
                                                            <attribute name='solutioncomponentid'/>
                                                            <link-entity name='pluginassembly' from='pluginassemblyid' to='objectid'>
                                                              <attribute name='pluginassemblyid'/>
                                                              <filter type='and'>
                                                                <condition attribute='name' operator='eq' value='{assemblyName}'/>
                                                              </filter>
                                                            </link-entity>
                                                            <link-entity name='solution' from='solutionid' to='solutionid'>
                                                              <attribute name='solutionid'/>
                                                              <filter type='and'>
                                                                <condition attribute='uniquename' operator='eq' value='{uniqueName}'/>
                                                              </filter>
                                                            </link-entity>
                                                          </entity>
                                                        </fetch>");

                EntityCollection results = client.RetrieveMultiple(query);

                return results.Entities.Count > 0;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error adding assembly to solution: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error adding assembly to solution: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return true;
            }
        }
    }
}