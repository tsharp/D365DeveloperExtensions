using System;

namespace SolutionPackager.ViewModels
{
    public class CrmSolution
    {
        public Guid SolutionId { get; set; }
        public string Name { get; set; }
        public string Prefix { get; set; }
        public string UniqueName { get; set; }
        public Version Version { get; set; }
        public string NameVersion { get; set; }
    }
}