using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NLog;
using Logger = NLog.Logger;

namespace CrmDeveloperExtensions2.Core
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
#pragma warning disable CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    public class UserOptionsGrid : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private bool _extensionLoggingEnabled;
        private bool _xrmToolingLoggingEnabled;

        public delegate void ApplySettingsHandler(object sender, EventArgs e); //Applied on save of form -may not be required
        public event PropertyChangedEventHandler PropertyChanged;

        //Need to add entry to CrmDeveloperExtensions2017 - CrmDeveloperExtensions2017Package.cs [ProvideOptionPage]

        [Category("Web Browser Options")]
        [DisplayName("Use internal VS web browser?")]
        [Description("Use the internal Visual Studio browser or your default browser for web content")]
        public bool UseInternalBrowser { get; set; } = false;

        [Category("Logging Options")]
        [DisplayName("Enable detailed extension logging?")]
        [Description("Detailed extension logging will log a bunch of stuff")]
        public bool ExtensionLoggingEnabled
        {
            get => _extensionLoggingEnabled;
            set
            {
                if (_extensionLoggingEnabled == value)
                    return;

                _extensionLoggingEnabled = value;
                NotifyPropertyChanged();
            }
        }

        //TODO: validate file path format
        [Category("Logging Options")]
        [DisplayName("Extension log file path")]
        [Description("Path to extension log file storage")]
        public string ExtensionLogFilePath { get; set; } = String.Empty;

        //TODO: message - requires reconnecting to CRM
        [Category("Logging Options")]
        [DisplayName("Enable Xrm.Tooling logging?")]
        [Description("Xrm.Tooling logging will log a bunch of stuff")]
        public bool XrmToolingLoggingEnabled
        {
            get => _xrmToolingLoggingEnabled;
            set
            {
                if (_xrmToolingLoggingEnabled == value)
                    return;

                _xrmToolingLoggingEnabled = value;
                NotifyPropertyChanged();
            }
        }

        //TODO: validate file path format
        [Category("Logging Options")]
        [DisplayName("Xrm.Tooling log file path")]
        [Description("Path to Xrm.Tooling log file storage")]
        public string XrmToolingLogFilePath { get; set; } = String.Empty;

        //TODO: validate file path format
        [Category("External Tools Options")]
        [DisplayName("Plug-in Registration Tool path")]
        [Description("Path to latest Plug-in Registration Tool executable")]
        public string PluginRegistrationToolPath { get; set; } = String.Empty;

        //TODO: validate file path format
        [Category("External Tools Options")]
        [DisplayName("Solution Packager Tool path")]
        [Description("Path to latest Solution Packager executable")]
        public string SolutionPackagerToolPath { get; set; } = String.Empty;

        //TODO: validate file path format
        [Category("External Tools Options")]
        [DisplayName("CrmSvcUtil Tool path")]
        [Description("Path to latest CrmSvcUtil executable")]
        public string CrmSvcUtilToolPath { get; set; } = String.Empty;

        [Category("Intellisense Options")]
        [DisplayName("Use Intellisense?")]
        [Description("Get metadata from CRM and provide string completion - this will require closing and re-opening any files")]
        public bool UseIntellisense { get; set; } = false;

        protected override void OnApply(PageApplyEventArgs e)
        {
            OnApplied(e);
            base.OnApply(e);
        }

        public event ApplySettingsHandler Applied;

        private void OnApplied(PageApplyEventArgs e)
        {
            Applied?.Invoke(this, e);
        }

        private void NotifyPropertyChanged([CallerMemberName] string value = null)
        {
            if (value == "ExtensionLoggingEnabled")
                ExtensionLogger.Log(LogLevel.Info, ExtensionLoggingEnabled ?
                    Resources.Resource.ExtensionLoggingEnabled :
                    Resources.Resource.ExtensionLoggingDisabled);

            if (value == "XrmToolingLoggingEnabled")
                ExtensionLogger.Log(LogLevel.Info, XrmToolingLoggingEnabled ?
                    Resources.Resource.XrmToolingLoggingEnabled :
                    Resources.Resource.XrmToolingLoggingDisabled);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }

        public static bool GetLoggingOptionBoolean(DTE dte, string propertyName)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsLoggingPage];

            return (bool)props.Item(propertyName).Value;
        }

        public static string GetLoggingOptionString(DTE dte, string propertyName)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsLoggingPage];

            return props.Item(propertyName).Value.ToString();
        }

        public static bool GetUseInternalBrowser(DTE dte)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsWebBrowserPage];

            return (bool)props.Item("UseInternalBrowser").Value;
        }

        public static string GetPluginRegistraionToolPath(DTE dte)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsToolsPage];

            return props.Item("PluginRegistrationToolPath").Value.ToString();
        }

        public static string GetSolutionPackagerToolPath(DTE dte)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsToolsPage];

            return props.Item("SolutionPackagerToolPath").Value.ToString();
        }

        public static string GetCrmSvcUtilToolPath(DTE dte)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsToolsPage];

            return props.Item("CrmSvcUtilToolPath").Value.ToString();
        }

        public static bool GetUseIntellisense(DTE dte)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsIntellisensePage];

            return (bool)props.Item("UseIntellisense").Value;
        }
    }
}