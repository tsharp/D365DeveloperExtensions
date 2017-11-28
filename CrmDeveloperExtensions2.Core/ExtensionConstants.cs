using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Models;
using System;
using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core
{
    public static class ExtensionConstants
    {
        public static Guid DefaultSolutionId = new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238");
        public static Guid UnitTestProjectType = new Guid("3AC096D0-A1C2-E12C-1390-A8335801FDAB");
        public static string IlMergeNuGet = "MSBuild.ILMerge.Task";

        public static string MicrosoftXrmSdk = "Microsoft.Xrm.Sdk";
        public static string MMicrosoftCrmSdkProxy = "Microsoft.Crm.Sdk.Proxy";
        public static string MicrosoftXrmSdkDeployment = "Microsoft.Xrm.Sdk.Deployment";
        public static string MicrosoftXrmClient = "Microsoft.Xrm.Client";
        public static string MicrosoftXrmPortal = "Microsoft.Xrm.Portal";
        public static string MicrosoftXrmSdkWorkflow = "Microsoft.Xrm.Sdk.Workflow";
        public static string MicrosoftXrmToolingConnector = "Microsoft.Xrm.Tooling.Connector";

        public static List<ToolWindow> CrmDevExToolWindows = new List<ToolWindow> {
            new ToolWindow {
                ToolWindowsId = new Guid("A3479AE0-5F4F-4A14-96F4-46F39000023A"),
                Type = ToolWindowType.WebResourceDeployer
            },
            new ToolWindow {
                ToolWindowsId = new Guid("F8BF1118-57B6-4404-9923-8A98AB710EBA"),
                Type = ToolWindowType.SolutionPackager
            },
            new ToolWindow {
                ToolWindowsId = new Guid("E7A15FDA-6C33-48F8-A1E7-D78E49458A7A"),
                Type = ToolWindowType.PluginTraceViewer
            },
            new ToolWindow {
                ToolWindowsId = new Guid("FA0E0759-D337-4C4C-8474-217A6BDC3C06"),
                Type = ToolWindowType.PluginDeployer
            }
        };

        public static string SolutionPackagerLogFile = "SolutionPackager.log";
        public static string SolutionPackagerMapFile = "packager_map.xml";
        public static string NoProfilesText = "Unnamed Profile";

        public static string SpklConfigFile = "spkl.json";
        public static string TemplateConfigFile = "templates.json";
        public static string DefaultPacakgeFolder = "package";
        public static string SpklRegAttrClassName = "CrmPluginRegistrationAttribute";

        public static string VsProjectTypeWebSite = "{E24C65DC-7377-472B-9ABA-BC803B73C61A}";
    }
}