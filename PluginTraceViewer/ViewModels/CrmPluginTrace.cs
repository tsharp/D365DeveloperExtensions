using System;

namespace PluginTraceViewer.ViewModels
{
    public class CrmPluginTrace
    {
        public string Entity { get; set; }
        public string CorrelationId { get; set; }
        public string MessageBlock { get; set; }
        public string MessageName { get; set; }
        public int Depth { get; set; }
        public int ExecutionDurationMs { get; set; }
        public string ExecutionDuration { get; set; }
        public string TypeName { get; set; }
        public DateTime CreatedOnUtc { get; set; }
        public DateTime CreatedOn { get; set; }
    }
}