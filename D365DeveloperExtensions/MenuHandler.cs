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
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace D365DeveloperExtensions
{
    public class MenuHandler : AsyncPackage
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package, DTE dte)
        {
            try
            {
                if (!(await package.GetServiceAsync((typeof(IMenuCommandService))) is OleMenuCommandService mcs))
                    throw new ArgumentNullException(Core.Resources.Resource.ErrorMessage_ErrorAccessingMCS);

                ExLogger.LogToFile(Logger, Resource.TraceInfo_InitializingMenu, LogLevel.Info);

                //Plug-in Deployer
                var pdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginDeployerWindow);
                var pdWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<PluginDeployerHost>(sender, e, package), pdWindowCommandId);
                mcs.AddCommand(pdWindowItem);

                //Web Resource Deployer
                var wrdWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
                var wrdWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<WebResourceDeployerHost>(sender, e, package), wrdWindowCommandId);
                mcs.AddCommand(wrdWindowItem);

                //Solution Packager
                var spWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidSolutionPackagerWindow);
                var spWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<SolutionPackagerHost>(sender, e, package), spWindowCommandId);
                mcs.AddCommand(spWindowItem);

                //Plug-in Trace Viewer
                var ptvWindowCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidPluginTraceViewerWindow);
                var ptvWindowItem = new OleMenuCommand((sender, e) => ShowToolWindow<PluginTraceViewerHost>(sender, e, package), ptvWindowCommandId);
                mcs.AddCommand(ptvWindowItem);

                //CRM Intellisense On
                var crmIntellisenseOnCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOn);
                var crmIntellisenseOnItem =
                    new OleMenuCommand((sender, e) => ToggleCrmIntellisense(sender, e, dte), crmIntellisenseOnCommandId)
                    {
                        Visible = false
                    };
                crmIntellisenseOnItem.BeforeQueryStatus += (sender, e) => DisplayCrmIntellisense(sender, e, dte);
                mcs.AddCommand(crmIntellisenseOnItem);

                //CRM Intellisense Off
                var crmIntellisenseOffCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidCrmIntellisenseOff);
                var crmIntellisenseOffItem =
                    new OleMenuCommand((sender, e) => ToggleCrmIntellisense(sender, e, dte), crmIntellisenseOffCommandId)
                    {
                        Visible = false
                    };
                crmIntellisenseOnItem.BeforeQueryStatus += (sender, e) => DisplayCrmIntellisense(sender, e, dte);
                mcs.AddCommand(crmIntellisenseOffItem);

                //NuGet SDK Tools - Core Tools
                var nugetSdkToolsCoreCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidNuGetSdkToolsCore);
                var nugetSdkToolsCoreItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsCoreCommandId);
                mcs.AddCommand(nugetSdkToolsCoreItem);

                //NuGet SDK Tools - Plug-in Registration Tool
                var nugetSdkToolsPrtCommandId = new CommandID(PackageGuids.GuidD365DevExCmdSet, PackageIds.CmdidNuGetSdkToolsPrt);
                var nugetSdkToolsPrtItem = new OleMenuCommand(InstallNuGetCliPackage, nugetSdkToolsPrtCommandId);
                mcs.AddCommand(nugetSdkToolsPrtItem);
            }
            catch (Exception e)
            {
                ExceptionHandler.LogException(Logger, null, e);
                throw;
            }
        }

        private static void DisplayCrmIntellisense(object sender, EventArgs eventArgs, DTE dte)
        {
            if (!(sender is OleMenuCommand menuCommand))
                return;

            var useIntellisense = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseIntellisense);
            if (!useIntellisense)
                return;

            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", dte);
            if (value == null)
            {
                menuCommand.Visible = menuCommand.CommandID.ID == 264;
                return;
            }

            var isEnabled = (bool)value;
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

            ExLogger.LogToFile(Logger, $"{Resource.Message_CRMIntellisenseEnabled}: {isEnabled}", LogLevel.Info);

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

                ExLogger.LogToFile(Logger, Resource.Message_ClearingMetadata, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_ClearingMetadata, MessageType.Info);

                return;
            }

            var result = MessageBox.Show(Resource.MessageBox_ConnectToCrm, Resource.MessageBox_ConnectToCrm_Title,
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);

            if (result != MessageBoxResult.Yes)
                return;

            ConnectToCrm();
        }

        private static void ConnectToCrm()
        {
            ExLogger.LogToFile(Logger, Resource.Message_OpeningCrmLoginForm, LogLevel.Info);

            var control = new CrmLoginForm(false);
            control.ConnectionToCrmCompleted += ControlOnConnectionToCrmCompleted;
            control.ShowDialog();
        }

        private static void ControlOnConnectionToCrmCompleted(object sender, EventArgs eventArgs)
        {
            ExLogger.LogToFile(Logger, Resource.Message_ClosingCrmLoginForm, LogLevel.Info);

            ((CrmLoginForm)sender).Close();
        }

        private static void InstallNuGetCliPackage(object sender, EventArgs e)
        {
            var oleMenuCommand = (OleMenuCommand)sender;

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
            ExLogger.LogToFile(Logger, $"{Resource.Message_OpeningToolWindow}: {typeof(T)}", LogLevel.Info);

            var window = package.FindToolWindow(typeof(T), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException(Resource.ErrorMessage_CannotCreateToolWindow);

            var windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
