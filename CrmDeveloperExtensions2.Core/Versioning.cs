using System;
using System.Text.RegularExpressions;

namespace CrmDeveloperExtensions2.Core
{
    public static class Versioning
    {
        public static Version StringToVersion(string version)
        {
            if (string.IsNullOrEmpty(version))
                return new Version(0, 0, 0, 0);

            string cleanVersion = Regex.Replace(version, "[^0-9.]", String.Empty);
            return Version.Parse(cleanVersion);
        }

        public static Version SolutionNameToVersion(string name)
        {
            name = name.ToLower().Replace(".zip", String.Empty);
            int index = name.IndexOf("_", StringComparison.Ordinal);
            name = name.Remove(index, 1);
            name = name.Replace("_", ".");
            name = name.Substring(1);

            return StringToVersion(name);
        }
    }
}