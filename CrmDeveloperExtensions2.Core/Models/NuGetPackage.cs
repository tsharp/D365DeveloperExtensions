using System;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class NuGetPackage
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public Version Version { get; set; }
        public string VersionText { get; set; }
    }
}
