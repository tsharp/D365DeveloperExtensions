using D365DeveloperExtensions.Core.Models;
using System;

namespace WebResourceDeployer.ViewModels
{
    public class NewWebResource
    {
        public Guid SolutionId { get; set; }
        public string File { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public WebResourceType WebResoruceType { get; set; }
        public string Description { get; set; }
    }
}