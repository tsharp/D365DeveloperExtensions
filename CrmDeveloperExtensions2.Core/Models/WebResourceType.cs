using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class WebResourceType
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public int Type { get; set; }
        public int CrmMinimumMajorVersion { get; set; }
        public int CrmMaximumMajorVersion { get; set; }
    }

    public static class WebResourceTypes
    {
        public static List<WebResourceType> Types => new List<WebResourceType>
        {
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "HTML", DisplayName = "Webpage (HTML)", Type = 1},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "CSS", DisplayName = "Style Sheet (CSS)", Type = 2},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "JS", DisplayName = "Script (JScript)", Type = 3},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "XML", DisplayName = "Data (XML)", Type = 4},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "PNG", DisplayName = "PNG format", Type = 5},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "JPG", DisplayName = "JPG format", Type = 6},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "GIF", DisplayName = "GIF format", Type = 7},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "XAP", DisplayName = "Silverlight (XAP)", Type = 8},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "XSL", DisplayName = "Style Sheet (XSL)", Type = 9},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99,Name = "ICO", DisplayName = "ICO format", Type = 10},
            new WebResourceType{CrmMinimumMajorVersion = 9, CrmMaximumMajorVersion = 99,Name = "SVG", DisplayName = "SVG format", Type = 11},
            new WebResourceType{CrmMinimumMajorVersion = 9, CrmMaximumMajorVersion = 99,Name = "RESX", DisplayName = "RESX format", Type = 12},
        };
    }
}