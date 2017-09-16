using System;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class CrmDevExAssembly
    {
        public Guid AssemblyId { get; set; }
        public Guid SolutionId { get; set; }
        public int DeploymentType { get; set; }
    }
}