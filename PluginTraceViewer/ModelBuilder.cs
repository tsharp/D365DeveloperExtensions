using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.Generic;

namespace PluginTraceViewer
{
    public static class ModelBuilder
    {
        public static List<CrmPluginTrace> CreateCrmPluginTraceView(EntityCollection pluginTraces)
        {
            List<CrmPluginTrace> crmPluginTraces = new List<CrmPluginTrace>();

            foreach (Entity pluginTrace in pluginTraces.Entities)
            {
                CrmPluginTrace crmPluginTrace = new CrmPluginTrace
                {
                    Entity = pluginTrace.GetAttributeValue<string>("primaryentity"),
                    CorrelationId = pluginTrace.GetAttributeValue<Guid>("correlationid").ToString(),
                    CreatedOn = pluginTrace.GetAttributeValue<DateTime>("createdon").ToLocalTime(),
                    CreatedOnUtc = pluginTrace.GetAttributeValue<DateTime>("createdon"),
                    Depth = pluginTrace.GetAttributeValue<int>("depth"),
                    ExecutionDurationMs = pluginTrace.GetAttributeValue<int>("performanceexecutionduration"),
                    ExecutionDuration = DateFormatting.MsToReadableTime(pluginTrace.GetAttributeValue<int>("performanceexecutionduration")),
                    MessageName = pluginTrace.GetAttributeValue<string>("messagename"),
                    MessageBlock = pluginTrace.GetAttributeValue<string>("messageblock"),
                    TypeName = pluginTrace.GetAttributeValue<string>("typename")
                };

                crmPluginTraces.Add(crmPluginTrace);
            }

            return crmPluginTraces;
        }
    }
}
