using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CrmDeveloperExtensions.Core
{
    public static class Versioning
    {
        public static Version StringToVersion(string version)
        {
            string cleanVersion = Regex.Replace(version, "[^0-9.]", "");
            return Version.Parse(cleanVersion);
        }
    }
}
