using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using CrmDeveloperExtensions2.Core;
using PluginDeployer;
using SolutionPackager;
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
    [InstalledProductRegistration("#110", "#112", "2.0.0.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(PluginDeployerHost))]
    [ProvideToolWindow(typeof(WebResourceDeployerHost))]
    [ProvideToolWindow(typeof(SolutionPackagerHost))]

    [Guid(PackageGuids.GuidCrmDeveloperExtensionsPkgString)]
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]

    //User Settings - Sections
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Logging", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "Web Browser", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGrid), "Crm DevEx", "External Tools", 0, 0, true)]

    public sealed class CrmDeveloperExtensions2Package : Package
    {
        private DTE _dte;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();

        protected override void Initialize()
        {
            base.Initialize();

            _dte = GetGlobalService(typeof(DTE)) as DTE;

            ExLogger.LogToFile(_dte, ExtensionLogger, "Initializing extension", LogLevel.Info);

            StartupTasks.Run(_dte);

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs == null) return;

            //Plug-in Deployer
            CommandID pdWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidPluginDeployerWindow);
            OleMenuCommand pdWindowItem = new OleMenuCommand(ShowPluginDeployer, pdWindowCommandId);
            mcs.AddCommand(pdWindowItem);

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand(ShowWebResourceDeployer, wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);

            //Solution Packager
            CommandID spWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidSolutionPackagerWindow);
            OleMenuCommand spWindowItem = new OleMenuCommand(ShowSolutionPackager, spWindowCommandId);
            mcs.AddCommand(spWindowItem);
        }

        private void ShowPluginDeployer(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(PluginDeployerHost), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException("Cannot create tool window.");

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void ShowWebResourceDeployer(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(WebResourceDeployerHost), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException("Cannot create tool window.");

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void ShowSolutionPackager(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(SolutionPackagerHost), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException("Cannot create tool window.");

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
