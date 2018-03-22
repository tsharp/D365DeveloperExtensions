using System.Collections.Generic;

namespace D365DeveloperExtensions.Core.Models
{
    public class NpmHistory
    {
        public string name { get; set; }
        public string description { get; set; }
        public string license { get; set; }
        public List<string> versions { get; set; }
    }
}