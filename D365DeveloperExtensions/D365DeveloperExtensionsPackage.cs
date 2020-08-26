using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
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
using System.Runtime.InteropServices;
using System.Threading;
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
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.20239.1842", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(PluginDeployerHost))]
    [ProvideToolWindow(typeof(WebResourceDeployerHost))]
    [ProvideToolWindow(typeof(SolutionPackagerHost))]
    [ProvideToolWindow(typeof(PluginTraceViewerHost))]

    [Guid(PackageGuids.GuidD365DeveloperExtensionsPkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.WizardOpen_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.FirstLaunchSetup_string, PackageAutoLoadFlags.BackgroundLoad)]

    //User Settings - Sections
    //TODO: find way to replace strings
    [ProvideOptionPage(typeof(UserOptionsGridLogging), "D365 DevEx", "Logging", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridWebBrowser), "D365 DevEx", "Web Browser", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridTools), "D365 DevEx", "External Tools", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridIntellisense), "D365 DevEx", "Intellisense", 0, 0, true)]
    [ProvideOptionPage(typeof(UserOptionsGridTemplates), "D365 DevEx", "Templates", 0, 0, true)]

    public sealed class D365DeveloperExtensionsPackage : AsyncPackage
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private IVsSolution _vsSolution;
        private IVsSolutionEvents _vsSolutionEvents;

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // Switches to the UI thread in order to consume some services used in command initialization
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            try
            {
                if (!(GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Core.Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                GetUserOptions(dte);

                StartupTasks.Run();

                ExLogger.LogToFile(Logger, Resource.TraceInfo_InitializingExtension, LogLevel.Info);

                AdviseSolutionEvents();
                var events = dte.Events;

                BindSolutionEvents(events);

                await MenuHandler.InitializeAsync(this, dte);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, null, ex);
                throw;
            }
        }

        private void AdviseSolutionEvents()
        {
            _vsSolutionEvents = new VsSolutionEvents(this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out _);
        }

        private void BindSolutionEvents(Events events)
        {
            var solutionEvents = events.SolutionEvents;
            solutionEvents.BeforeClosing += SolutionEventsOnBeforeClosing;
        }

        public void SolutionEventsOnBeforeClosing()
        {
            try
            {
                if (!(GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Core.Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                SharedGlobals.SetGlobal("UseCrmIntellisense", null, dte);

                if (SharedGlobals.GetGlobal("CrmService", dte) != null)
                    SharedGlobals.SetGlobal("CrmService", null, dte);

                if (SharedGlobals.GetGlobal("CrmMetadata", dte) != null)
                {
                    SharedGlobals.SetGlobal("CrmMetadata", null, dte);
                    OutputLogger.WriteToOutputWindow("Clearing metadata", MessageType.Info);
                }
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, null, ex);
                throw;
            }
        }

        private void GetUserOptions(DTE dte)
        {
            var intellisenseOptions = (UserOptionsGridIntellisense)GetDialogPage(typeof(UserOptionsGridIntellisense));
            SharedGlobals.SetGlobal("IntellisenseEntityTriggerCharacter", intellisenseOptions.IntellisenseEntityTriggerCharacter, dte);
            SharedGlobals.SetGlobal("IntellisenseFieldTriggerCharacter", intellisenseOptions.IntellisenseFieldTriggerCharacter, dte);
            SharedGlobals.SetGlobal("UseIntellisense", intellisenseOptions.UseIntellisense, dte);

            var loggingOptions = (UserOptionsGridLogging)GetDialogPage(typeof(UserOptionsGridLogging));
            SharedGlobals.SetGlobal("ExtensionLoggingEnabled", loggingOptions.ExtensionLoggingEnabled, dte);
            SharedGlobals.SetGlobal("ExtensionLogFilePath", loggingOptions.ExtensionLogFilePath, dte);
            SharedGlobals.SetGlobal("XrmToolingLogFilePath", loggingOptions.XrmToolingLogFilePath, dte);
            SharedGlobals.SetGlobal("XrmToolingLoggingEnabled", loggingOptions.XrmToolingLoggingEnabled, dte);

            var templateOptions = (UserOptionsGridTemplates)GetDialogPage(typeof(UserOptionsGridTemplates));
            SharedGlobals.SetGlobal("CustomTemplatesPath", templateOptions.CustomTemplatesPath, dte);
            SharedGlobals.SetGlobal("DefaultKeyFileName", templateOptions.DefaultKeyFileName, dte);

            var toolsOptions = (UserOptionsGridTools)GetDialogPage(typeof(UserOptionsGridTools));
            SharedGlobals.SetGlobal("CrmSvcUtilToolPath", toolsOptions.CrmSvcUtilToolPath, dte);
            SharedGlobals.SetGlobal("PluginRegistrationToolPath", toolsOptions.PluginRegistrationToolPath, dte);
            SharedGlobals.SetGlobal("SolutionPackagerToolPath", toolsOptions.SolutionPackagerToolPath, dte);

            var webBrowserOptions = (UserOptionsGridWebBrowser)GetDialogPage(typeof(UserOptionsGridWebBrowser));
            SharedGlobals.SetGlobal("UseInternalBrowser", webBrowserOptions.UseInternalBrowser, dte);
        }
    }
}