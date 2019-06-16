using CrmIntellisense.Crm;
using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using D365DeveloperExtensions.Resources;
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
using System.Windows;
using TemplateWizards;
using WebResourceDeployer;

namespace D365DeveloperExtensions
{
    public class MenuHandler : AsyncPackage
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package, DTE dte)
        {
            var mcs = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;

            ExtensionLogger.LogToFile(Logger, Resource.TraceInfo_InitializingExtension, LogLevel.Info);

            //Plug-in Deployer
            CommandID pdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginDeployerWindow);
            OleMenuCommand pdWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<PluginDeployerHost>(sender, e, package), pdWindowCommandId);
            mcs.AddCommand(pdWindowItem);

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<WebResourceDeployerHost>(sender, e, package), wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);

            //Solution Packager
            CommandID spWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidSolutionPackagerWindow);
            OleMenuCommand spWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<SolutionPackagerHost>(sender, e, package), spWindowCommandId);
            mcs.AddCommand(spWindowItem);

            //Plug-in Trace Viewer
            CommandID ptvWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginTraceViewerWindow);
            OleMenuCommand ptvWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<PluginTraceViewerHost>(sender, e, package), ptvWindowCommandId);
            mcs.AddCommand(ptvWindowItem);

            //CRM Intellisense On
            CommandID crmIntellisenseOnCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOn);
            OleMenuCommand crmIntellisenseOnItem =
                new OleMenuCommand((sender, e) => ToggleCrmIntellisense(sender, e, dte), crmIntellisenseOnCommandId)
                {
                    Visible = false
                };
            crmIntellisenseOnItem.BeforeQueryStatus += (sender, e) => DisplayCrmIntellisense(sender, e, dte);
            mcs.AddCommand(crmIntellisenseOnItem);

            //CRM Intellisense Off
            CommandID crmIntellisenseOffCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOff);
            OleMenuCommand crmIntellisenseOffItem =
                new OleMenuCommand((sender, e) => ToggleCrmIntellisense(sender, e, dte), crmIntellisenseOffCommandId)
                {
                    Visible = false
                };
            crmIntellisenseOnItem.BeforeQueryStatus += (sender, e) => DisplayCrmIntellisense(sender, e, dte);
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

        private static void DisplayCrmIntellisense(object sender, EventArgs eventArgs, DTE dte)
        {
            if (!(sender is OleMenuCommand menuCommand))
                return;

            bool useIntellisense = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseIntellisense);
            if (!useIntellisense)
                return;

            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", dte);
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

        private static void ToggleCrmIntellisense(object sender, EventArgs e, DTE dte)
        {
            bool isEnabled;
            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", dte);
            if (value == null)
            {
                isEnabled = false;
                SharedGlobals.SetGlobal("UseCrmIntellisense", true, dte);
            }
            else
            {
                isEnabled = (bool)value;
                SharedGlobals.SetGlobal("UseCrmIntellisense", !isEnabled, dte);
            }

            if (!isEnabled) //On
            {
                if (HostWindow.IsCrmDexWindowOpen(dte) && SharedGlobals.GetGlobal("CrmService", dte) != null)
                    return;
            }
            else
            {
                if (!HostWindow.IsCrmDexWindowOpen(dte) && SharedGlobals.GetGlobal("CrmService", dte) != null)
                    SharedGlobals.SetGlobal("CrmService", null, dte);

                CrmMetadata.Metadata = null;
                SharedGlobals.SetGlobal("CrmMetadata", null, dte);
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

        private static void ShowToolWindow<T>(object sender, EventArgs e, AsyncPackage package)
        {
            ToolWindowPane window = package.FindToolWindow(typeof(T), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException(Resource.ErrorMessage_CannotCreateToolWindow);

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
