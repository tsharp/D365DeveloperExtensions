using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmDeveloperExtensions.Core.Models
{
    public class CrmDevExConfigOrgMap
    {
        public Guid OrganizationId { get; set; }
        public List<CrmDexExConfigWebResource> WebResources { get; set; }
    }
}