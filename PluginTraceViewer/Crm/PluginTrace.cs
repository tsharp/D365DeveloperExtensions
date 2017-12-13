using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using PluginTraceViewer.Resources;
using System;
using System.Collections.Generic;

namespace PluginTraceViewer.Crm
{
    public static class PluginTrace
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

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
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingTraces, ex);

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

                CheckForDeletionErrors(pluginTraceLogIds, executeMultipleResponse, deletedPluginTraceLogIds);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorDeletingTraces, ex);
            }

            return deletedPluginTraceLogIds;
        }

        private static void CheckForDeletionErrors(Guid[] pluginTraceLogIds, ExecuteMultipleResponse executeMultipleResponse,
            List<Guid> deletedPluginTraceLogIds)
        {
            foreach (var responseItem in executeMultipleResponse.Responses)
            {
                if (responseItem.Response != null)
                {
                    deletedPluginTraceLogIds.Add(pluginTraceLogIds[responseItem.RequestIndex]);
                    continue;
                }

                if (responseItem.Fault != null)
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_ErrorDeletingTrace}: {responseItem.Fault}", MessageType.Error);
            }
        }
    }
}