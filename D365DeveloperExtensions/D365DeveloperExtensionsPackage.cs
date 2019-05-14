using CrmIntellisense.Crm;
using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using D365DeveloperExtensions.Resources;
using D365DeveloperExtensions.Vs;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using PluginDeployer;
using PluginTraceViewer;
using SolutionPackager;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Windows;
using TemplateWizards;
using WebResourceDeployer;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace D365DeveloperExtensions
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
    [InstalledProductRegistration("#110", "#112", "2.0.19134.1636", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(PluginDeployerHost))]
    [ProvideToolWindow(typeof(WebResourceDeployerHost))]
    [ProvideToolWindow(typeof(SolutionPackagerHost))]
    [ProvideToolWindow(typeof(PluginTraceViewerHost))]

    [Guid(PackageGuids.GuidD365DeveloperExtensionsPkgString)]
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]

    //User Settings - Sections
    //TODO: find way to replace strings
    [ProvideOptionPage(typeof(UserOptionsGridLogging), "D365 DevEx", "Logging", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridWebBrowser), "D365 DevEx", "Web Browser", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridTools), "D365 DevEx", "External Tools", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridIntellisense), "D365 DevEx", "Intellisense", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridTemplates), "D365 DevEx", "Templates", 0, 0, true)]

    public sealed class D365DeveloperExtensionsPackage : Package
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static DTE _dte;
        private IVsSolution _vsSolution;
        private IVsSolutionEvents _vsSolutionEvents;

        protected override void Initialize()
        {
            base.Initialize();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            StartupTasks.Run(_dte);

            ExLogger.LogToFile(Logger, Resource.TraceInfo_InitializingExtension, LogLevel.Info);

            AdviseSolutionEvents();
            var events = _dte.Events;

            BindSolutionEvents(events);

            if (!(GetService(typeof(IMenuCommandService)) is OleMenuCommandService mcs))
                return;

            //Plug-in Deployer
            CommandID pdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginDeployerWindow);
            OleMenuCommand pdWindowItem = new OleMenuCommand(ShowToolWindow<PluginDeployerHost>, pdWindowCommandId);
            mcs.AddCommand(pdWindowItem);

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand(ShowToolWindow<WebResourceDeployerHost>, wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);

            //Solution Packager
            CommandID spWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidSolutionPackagerWindow);
            OleMenuCommand spWindowItem = new OleMenuCommand(ShowToolWindow<SolutionPackagerHost>, spWindowCommandId);
            mcs.AddCommand(spWindowItem);

            //Plug-in Trace Viewer
            CommandID ptvWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginTraceViewerWindow);
            OleMenuCommand ptvWindowItem = new OleMenuCommand(ShowToolWindow<PluginTraceViewerHost>, ptvWindowCommandId);
            mcs.AddCommand(ptvWindowItem);

            //CRM Intellisense On
            CommandID crmIntellisenseOnCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOn);
            OleMenuCommand crmIntellisenseOnItem =
                new OleMenuCommand(ToggleCrmIntellisense, crmIntellisenseOnCommandId)
                {
                    Visible = false
                };
            crmIntellisenseOnItem.BeforeQueryStatus += DisplayCrmIntellisense;
            mcs.AddCommand(crmIntellisenseOnItem);

            //CRM Intellisense Off
            CommandID crmIntellisenseOffCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOff);
            OleMenuCommand crmIntellisenseOffItem =
                new OleMenuCommand(ToggleCrmIntellisense, crmIntellisenseOffCommandId)
                {
                    Visible = false
                };
            crmIntellisenseOffItem.BeforeQueryStatus += DisplayCrmIntellisense;
            mcs.AddCommand(crmIntellisenseOffItem);

            //NuGet SDK Tools - Core Tools
            CommandID nugetSdkToolsCoreCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidNuGetSdkToolsCore);
            OleMenuCommand nugetSdkToolsCoreItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsCoreCommandId);
            mcs.AddCommand(nugetSdkToolsCoreItem);

            //NuGet SDK Tools - Plug-in Registration Tool
            CommandID nugetSdkToolsPrtCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidNuGetSdkToolsPrt);
            OleMenuCommand nugetSdkToolsPrtItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsPrtCommandId);
            mcs.AddCommand(nugetSdkToolsPrtItem);
        }

        private static void DisplayCrmIntellisense(object sender, EventArgs eventArgs)
        {
            if (!(sender is OleMenuCommand menuCommand))
                return;

            bool useIntellisense = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseIntellisense);
            if (!useIntellisense)
                return;

            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", _dte);
            if (value == null)
            {
                menuCommand.Visible = menuCommand.CommandID.ID == 264;
                return;
            }

            bool isEnabled = (bool)value;
            if (isEnabled)
                menuCommand.Visible = menuCommand.CommandID.ID != 264;
            else
                menuCommand.Visible = menuCommand.CommandID.ID == 264;
        }

        private static void ToggleCrmIntellisense(object sender, EventArgs e)
        {
            bool isEnabled;
            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", _dte);
            if (value == null)
            {
                isEnabled = false;
                SharedGlobals.SetGlobal("UseCrmIntellisense", true, _dte);
            }
            else
            {
                isEnabled = (bool)value;
                SharedGlobals.SetGlobal("UseCrmIntellisense", !isEnabled, _dte);
            }

            if (!isEnabled) //On
            {
                if (HostWindow.IsCrmDexWindowOpen(_dte) && SharedGlobals.GetGlobal("CrmService", _dte) != null)
                    return;
            }
            else
            {
                if (!HostWindow.IsCrmDexWindowOpen(_dte) && SharedGlobals.GetGlobal("CrmService", _dte) != null)
                    SharedGlobals.SetGlobal("CrmService", null, _dte);

                CrmMetadata.Metadata = null;
                SharedGlobals.SetGlobal("CrmMetadata", null, _dte);
                OutputLogger.WriteToOutputWindow("Clearing metadata", MessageType.Info);

                return;
            }

            MessageBoxResult result = MessageBox.Show(Resource.MessageBox_ConnectToCrm, Resource.MessageBox_ConnectToCrm_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);

            if (result != MessageBoxResult.Yes)
                return;

            ConnectToCrm();
        }

        private static void ConnectToCrm()
        {
            CrmLoginForm ctrl = new CrmLoginForm(false);
            ctrl.ConnectionToCrmCompleted += CtrlOnConnectionToCrmCompleted;
            ctrl.ShowDialog();
        }

        private static void CtrlOnConnectionToCrmCompleted(object sender, EventArgs eventArgs)
        {
            ((CrmLoginForm)sender).Close();
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

        private void AdviseSolutionEvents()
        {
            _vsSolutionEvents = new VsSolutionEvents(_dte, this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out uint solutionEventsCookie);
        }

        private void BindSolutionEvents(Events events)
        {
            var solutionEvents = events.SolutionEvents;
            solutionEvents.BeforeClosing += SolutionEventsOnBeforeClosing;
        }

        public void SolutionEventsOnBeforeClosing()
        {
            SharedGlobals.SetGlobal("UseCrmIntellisense", null, _dte);

            if (SharedGlobals.GetGlobal("CrmService", _dte) != null)
                SharedGlobals.SetGlobal("CrmService", null, _dte);

            if (SharedGlobals.GetGlobal("CrmMetadata", _dte) != null)
            {
                SharedGlobals.SetGlobal("CrmMetadata", null, _dte);
                OutputLogger.WriteToOutputWindow("Clearing metadata", MessageType.Info);
            }
        }
    }
}