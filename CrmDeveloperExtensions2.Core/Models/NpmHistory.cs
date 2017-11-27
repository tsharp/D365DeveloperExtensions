using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class NpmHistory
    {
        public string name { get; set; }
        public string description { get; set; }
        public List<string> versions { get; set; }
    }
}