using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrmDeveloperExtensions2.Core.Models;

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