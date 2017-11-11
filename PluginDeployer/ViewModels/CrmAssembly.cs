using PluginDeployer.Spkl;
using System;

namespace PluginDeployer.ViewModels
{
    public class CrmAssembly
    {
        public Guid AssemblyId { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Version { get; set; }
        public bool IsWorkflow { get; set; }
        public Guid SolutionId { get; set; }
        public string Culture { get; set; }
        public string AssemblyPath { get; set; }
        public string PublicKeyToken { get; set; }
        public IsolationModeEnum IsolationMode { get; set; }
    }
}