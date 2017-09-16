using System;
using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class CrmDevExConfigOrgMap
    {
        public Guid OrganizationId { get; set; }
        public string ProjectUniqueName { get; set; }
        public List<CrmDexExConfigWebResource> WebResources { get; set; }
        public CrmDevExSolutionPackage SolutionPackage { get; set; }
        public CrmDevExAssembly Assembly { get; set; }
    }
}