using System;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class CrmDevExSolutionPackage
    {
        public Guid SolutionId { get; set; }
        public bool SaveSolutions { get; set; }
        public string ProjectFolder { get; set; }
        public bool DownloadManaged { get; set; }
        public bool CreateManaged { get; set; }
        public bool EnableSolutionPackagerLog { get; set; }
        public bool PublishAll { get; set; }
    }
}