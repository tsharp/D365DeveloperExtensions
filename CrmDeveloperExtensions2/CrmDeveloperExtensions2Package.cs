using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.UserOptions;
using CrmDeveloperExtensions2.Resources;
//using CrmIntellisense;
using PluginDeployer;
using PluginTraceViewer;
using SolutionPackager;
using TemplateWizards;
using WebResourceDeployer;
using ExLogger = CrmDeveloperExtensions2.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace CrmDeveloperExtensions2
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.17362.2149", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(PluginDeployerHost))]
    [ProvideToolWindow(typeof(WebResourceDeployerHost))]
    [ProvideToolWindow(typeof(SolutionPackagerHost))]
    [ProvideToolWindow(typeof(PluginTraceViewerHost))]

    [Guid(PackageGuids.GuidCrmDeveloperExtensionsPkgString)]
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]

    //User Settings - Sections
    //TODO: find way to replace strings
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Logging", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Web Browser", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "External Tools", 0, 0, true)]
    //[ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Intellisense", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Templates", 0, 0, true)]

    public sealed class CrmDeveloperExtensions2Package : Package
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        protected override void Initialize()
        {
            base.Initialize();

            DTE dte = GetGlobalService(typeof(DTE)) as DTE;
            StartupTasks.Run(dte);

            ExLogger.LogToFile(Logger, Resource.TraceInfo_InitializingExtension, LogLevel.Info);

            if (!(GetService(typeof(IMenuCommandService)) is OleMenuCommandService mcs))
                return;

            //Plug-in Deployer
            CommandID pdWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidPluginDeployerWindow);
            OleMenuCommand pdWindowItem = new OleMenuCommand(ShowToolWindow<PluginDeployerHost>, pdWindowCommandId);
            mcs.AddCommand(pdWindowItem);

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand(ShowToolWindow<WebResourceDeployerHost>, wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);

            //Solution Packager
            CommandID spWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidSolutionPackagerWindow);
            OleMenuCommand spWindowItem = new OleMenuCommand(ShowToolWindow<SolutionPackagerHost>, spWindowCommandId);
            mcs.AddCommand(spWindowItem);

            //Plug-in Trace Viewer
            CommandID ptvWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidPluginTraceViewerWindow);
            OleMenuCommand ptvWindowItem = new OleMenuCommand(ShowToolWindow<PluginTraceViewerHost>, ptvWindowCommandId);
            mcs.AddCommand(ptvWindowItem);

            //NuGet SDK Tools - Core Tools
            CommandID nugetSdkToolsCoreCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidNuGetSdkToolsCore);
            OleMenuCommand nugetSdkToolsCoreItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsCoreCommandId);
            mcs.AddCommand(nugetSdkToolsCoreItem);

            //NuGet SDK Tools - Plug-in Registration Tool
            CommandID nugetSdkToolsPrtCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidNuGetSdkToolsPrt);
            OleMenuCommand nugetSdkToolsPrtItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsPrtCommandId);
            mcs.AddCommand(nugetSdkToolsPrtItem);
        }

        private static void InstallNuGetCliPackage(object sender, EventArgs e)
        {
            OleMenuCommand oleMenuCommand = (OleMenuCommand)sender;

            switch (oleMenuCommand.CommandID.ID)
            {
                case 262:
                    SdkToolsInstaller.InstallNuGetCliPackage(ExtensionConstants.MicrosoftCrmSdkXrmToolingPrt);
                    break;
                case 263:
                    SdkToolsInstaller.InstallNuGetCliPackage(ExtensionConstants.MicrosoftCrmSdkCoreTools);
                    break;
            }
        }

        private void ShowToolWindow<T>(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(T), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException(Resource.ErrorMessage_CannotCreateToolWindow);

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}