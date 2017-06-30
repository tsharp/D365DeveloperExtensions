using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;

namespace PluginTraceViewer
{
    public static class ModelBuilder
    {
        public static ObservableCollection<CrmPluginTrace> CreateCrmPluginTraceView(EntityCollection pluginTraces)
        {
            ObservableCollection<CrmPluginTrace> crmPluginTraces = new ObservableCollection<CrmPluginTrace>();

            foreach (Entity pluginTrace in pluginTraces.Entities)
            {
                CrmPluginTrace crmPluginTrace = new CrmPluginTrace
                {
                    PluginTraceLogidId = pluginTrace.Id,
                    Entity = pluginTrace.GetAttributeValue<string>("primaryentity"),
                    CorrelationId = pluginTrace.GetAttributeValue<Guid>("correlationid").ToString(),
                    CreatedOn = pluginTrace.GetAttributeValue<DateTime>("createdon").ToLocalTime(),
                    CreatedOnUtc = pluginTrace.GetAttributeValue<DateTime>("createdon"),
                    Depth = pluginTrace.GetAttributeValue<int>("depth"),
                    ExecutionDurationMs = pluginTrace.GetAttributeValue<int>("performanceexecutionduration"),
                    ExecutionDuration = DateFormatting.MsToReadableTime(pluginTrace.GetAttributeValue<int>("performanceexecutionduration")),
                    MessageName = pluginTrace.GetAttributeValue<string>("messagename"),
                    MessageBlock = pluginTrace.GetAttributeValue<string>("messageblock"),
                    TypeName = pluginTrace.GetAttributeValue<string>("typename"),
                    Mode = (pluginTrace.GetAttributeValue<OptionSetValue>("mode").Value == 0) ? "Synchronous" : "Asynchronous",
                    ExceptionDetails = pluginTrace.GetAttributeValue<string>("exceptiondetails"),
                    Details = CreateDetails(pluginTrace.GetAttributeValue<string>("messageblock"), pluginTrace.GetAttributeValue<string>("exceptiondetails"))
                };

                crmPluginTraces.Add(crmPluginTrace);
            }

            return crmPluginTraces;
        }

        private static string CreateDetails(string messageBlock, string exceptionDetails)
        {
            string result = String.Empty;

            if (!string.IsNullOrEmpty(messageBlock))
                result += messageBlock;

            if (!string.IsNullOrEmpty(result) && !string.IsNullOrEmpty(exceptionDetails))
                result += Environment.NewLine + Environment.NewLine;

            if (!string.IsNullOrEmpty(exceptionDetails))
                result += exceptionDetails;

            return result;
        }
    }
}