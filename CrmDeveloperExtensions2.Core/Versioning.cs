using System;
using System.Text.RegularExpressions;

namespace CrmDeveloperExtensions2.Core
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
