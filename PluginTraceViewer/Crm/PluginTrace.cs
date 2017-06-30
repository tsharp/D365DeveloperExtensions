using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Messages;

namespace PluginTraceViewer.Crm
{
    public static class PluginTrace
    {
        public static EntityCollection RetrievePluginTracesFromCrm(CrmServiceClient client, DateTime afterDate)
        {
            try
            {
                FetchExpression query = new FetchExpression($@"<fetch>
                                                                <entity name='plugintracelog' >
                                                                <attribute name='messagename' />
                                                                <attribute name='plugintracelogid' />
                                                                <attribute name='primaryentity' />
                                                                <attribute name='exceptiondetails' />
                                                                <attribute name='messageblock' />
                                                                <attribute name='performanceexecutionduration' />
                                                                <attribute name='createdon' />
                                                                <attribute name='typename' />
                                                                <attribute name='depth' />
                                                                <attribute name='mode' />
                                                                <attribute name='correlationid' />
                                                                <filter type='and' >
                                                                    <condition attribute='createdon' operator='gt' value='{afterDate}' />
                                                                </filter>
                                                                <order attribute='createdon' descending='true' />
                                                                </entity>
                                                            </fetch>");

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

        public static List<Guid> DeletePluginTracesFromCrm(CrmServiceClient client, Guid[] pluginTraceLogIds)
        {
            List<Guid> deletedPluginTraceLogIds = new List<Guid>();

            try
            {
                ExecuteMultipleRequest executeMultipleRequest = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    }
                };

                foreach (Guid pluginTraceLogId in pluginTraceLogIds)
                {
                    DeleteRequest request = new DeleteRequest
                    {
                        Target = new EntityReference("plugintracelog", pluginTraceLogId)
                    };

                    executeMultipleRequest.Requests.Add(request);
                }

                ExecuteMultipleResponse executeMultipleResponse =
                    (ExecuteMultipleResponse)client.Execute(executeMultipleRequest);

                foreach (var responseItem in executeMultipleResponse.Responses)
                {
                    if (responseItem.Response != null)
                    {
                        deletedPluginTraceLogIds.Add(pluginTraceLogIds[responseItem.RequestIndex]);
                        continue;
                    }

                    if (responseItem.Fault != null)
                        OutputLogger.WriteToOutputWindow(
                            "Error Deleting Plug-in Trace Log From CRM: " + responseItem.Fault, MessageType.Error);
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Deleting Plug-in Trace Log(s) From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Deleting Plug-in Trace Log(s) From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
            }

            return deletedPluginTraceLogIds;
        }
    }
}