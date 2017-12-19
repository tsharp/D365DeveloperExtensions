using System;

namespace WebResourceDeployer.ViewModels
{
    public class CrmSolution
    {
        public Guid SolutionId { get; set; }
        public string Name { get; set; }
        public string Prefix { get; set; }
        public string UniqueName { get; set; }
    }
}