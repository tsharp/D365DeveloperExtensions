using System;

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

        public static string[] CrmDevExWindows = { "A3479AE0-5F4F-4A14-96F4-46F39000023A", "FA0E0759-D337-4C4C-8474-217A6BDC3C06",
            "F8BF1118-57B6-4404-9923-8A98AB710EBA", "E7A15FDA-6C33-48F8-A1E7-D78E49458A7A" };

        public static string SolutionPackagerLogFile = "SolutionPackager.log";
        public static string SolutionPackagerMapFile = "mapping.xml";
    }
}