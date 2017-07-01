using System;

namespace CrmDeveloperExtensions2
{
    /// <summary>
    /// Helper class that exposes all GUIDs used across VS Package.
    /// </summary>
    internal sealed partial class PackageGuids
    {
        public const string GuidCrmDeveloperExtensionsPkgString = "A4FA8F78-98AF-4633-8621-BBDD7792A6AC";
        public const string GuidCrmDevExCmdSetString = "95CD7B0B-0592-4683-B42C-A79A41380FFE";
        public static Guid GuidCrmDeveloperExtensionsPkg = new Guid(GuidCrmDeveloperExtensionsPkgString);
        public static Guid GuidCrmDevExCmdSet = new Guid(GuidCrmDevExCmdSetString);
        public const string GuidCrmIntellisensePkgString = "C999FB53-B0D4-4AA6-8454-621D305A7CAF";
    }
    /// <summary>
    /// Helper class that encapsulates all CommandIDs uses across VS Package.
    /// </summary>
    internal sealed partial class PackageIds
    {
        public const int ProjectMenuGroup = 0x1020;
        public const int CmdidPluginDeployerWindow = 0x0101;
        public const int CmdidWebResourceDeployerWindow = 0x0102;
        public const int CmdidPluginTraceViewerWindow = 0x0103;
        public const int CmdidSolutionPackagerWindow = 0x0104;
        public const int CmdidCrmIntellisense = 0x0105;
        public const int TopLevelMenu = 0x0100;
        public const int TopLevelMenuGroup = 0x0200;
    }
}