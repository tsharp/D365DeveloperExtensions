using System;

namespace PluginDeployer.ViewModels
{
    public class CrmAssembly
    {
        public Guid AssemblyId { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public Version Version { get; set; }
        public bool IsWorkflow { get; set; }
        public Guid SolutionId { get; set; }
    }
}