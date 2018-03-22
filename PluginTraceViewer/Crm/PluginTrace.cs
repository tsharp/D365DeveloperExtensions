using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
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
            OutputLogger.WriteToOutputWindow(Resource.Message_RetrievingTraces, MessageType.Info);

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
                                                                    <condition attribute='createdon' operator='gt' value='{afterDate:o}' />
                                                                </filter>
                                                                <order attribute='createdon' descending='true' />
                                                                </entity>
                                                               </fetch>");

                EntityCollection traceLogs = client.RetrieveMultiple(query);

                if (traceLogs.Entities.Count > 0)
                    OutputLogger.WriteToOutputWindow($"{Resource.Info_RetrievedNewTraces}: " + traceLogs.Entities.Count, MessageType.Info);
                else
                    OutputLogger.WriteToOutputWindow(Resource.Info_NoNewTraces, MessageType.Info);

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
            OutputLogger.WriteToOutputWindow(Resource.Message_DeletingTraces, MessageType.Info);

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

                deletedPluginTraceLogIds = CheckForDeletionErrors(pluginTraceLogIds, executeMultipleResponse);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorDeletingTraces, ex);
            }

            return deletedPluginTraceLogIds;
        }

        private static List<Guid> CheckForDeletionErrors(Guid[] pluginTraceLogIds, ExecuteMultipleResponse executeMultipleResponse)
        {
            List<Guid> deletedPluginTraceLogIds = new List<Guid>();

            foreach (var responseItem in executeMultipleResponse.Responses)
            {
                if (responseItem.Response != null)
                {
                    deletedPluginTraceLogIds.Add(pluginTraceLogIds[responseItem.RequestIndex]);
                    OutputLogger.WriteToOutputWindow($"{Resource.Message_DeletedTraceLog}: {pluginTraceLogIds[responseItem.RequestIndex]}", MessageType.Info);
                    continue;
                }

                if (responseItem.Fault != null)
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_ErrorDeletingTrace}: {responseItem.Fault}", MessageType.Error);
            }

            return deletedPluginTraceLogIds;
        }
    }
}