using System;

namespace D365DeveloperExtensions.Core.Models
{
    public class NuGetPackage
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public Version Version { get; set; }
        public string VersionText { get; set; }
        public bool XrmToolingClient { get; set; }
        public string LicenseUrl { get; set; }
    }
}