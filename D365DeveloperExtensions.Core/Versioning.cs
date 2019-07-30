using System;
using System.Text.RegularExpressions;

namespace D365DeveloperExtensions.Core
{
    public static class Versioning
    {
        public static Version StringToVersion(string version)
        {
            if (string.IsNullOrEmpty(version))
                return new Version(0, 0, 0, 0);

            var cleanVersion = Regex.Replace(version, "[^0-9.]", string.Empty);

            return Version.Parse(cleanVersion);
        }

        public static Version SolutionNameToVersion(string name)
        {
            name = name.ToLower().Replace(".zip", string.Empty);
            var index = name.IndexOf("_", StringComparison.Ordinal);
            name = name.Remove(index, 1);
            name = name.Replace("_", ".");
            name = name.Substring(1);

            return StringToVersion(name);
        }

        public static bool DoAssemblyVersionsMatch(Version aVersion, Version bVersion)
        {
            return aVersion.Major == bVersion.Major &&
                   aVersion.Minor == bVersion.Minor;
        }

        public static Version ValidateVersionInput(string majorIn, string minorIn, string buildIn, string revisionIn)
        {
            var isMajorInt = int.TryParse(majorIn, out var major);
            var isMinorInt = int.TryParse(minorIn, out var minor);
            var isBuildInt = int.TryParse(buildIn, out var build);
            var isRevisionInt = int.TryParse(revisionIn, out var revision);

            if (!isMajorInt || !isMinorInt)
                return null;

            var v = string.Concat(major, ".", minor, isBuildInt ? $".{build}" : null,
                isRevisionInt ? $".{revision}" : null);

            return new Version(v);
        }
    }
}