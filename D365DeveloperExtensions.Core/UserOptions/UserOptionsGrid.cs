using D365DeveloperExtensions.Core.Localization;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using Logger = NLog.Logger;

namespace D365DeveloperExtensions.Core.UserOptions
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

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]

        [LocalizedCategory("UserOptions_Category_WebBrowser", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_WebBrowser", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_WebBrowser", typeof(Resource))]
        public bool UseInternalBrowser { get; set; } = false;

        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_ExtensionLogging", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_ExtensionLogging", typeof(Resource))]
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
        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_ExtensionLoggingPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_ExtensionLoggingPath", typeof(Resource))]
        public string ExtensionLogFilePath { get; set; } = String.Empty;

        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_XrmToolingLogging", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_XrmToolingLogging", typeof(Resource))]
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
        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_XrmToolingLoggingPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_XrmToolingLoggingPath", typeof(Resource))]
        public string XrmToolingLogFilePath { get; set; } = String.Empty;

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_RegToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_RegToolPath", typeof(Resource))]
        public string PluginRegistrationToolPath { get; set; } = String.Empty;

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_SolutionPackToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_SolutionPackToolPath", typeof(Resource))]
        public string SolutionPackagerToolPath { get; set; } = String.Empty;

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_CrmSvcToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_CrmSvcToolPath", typeof(Resource))]
        public string CrmSvcUtilToolPath { get; set; } = String.Empty;

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Templates", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_TemplatesPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_TemplatesPath", typeof(Resource))]
        public string CustomTemplatesPath { get; set; } = String.Empty;

        [LocalizedCategory("UserOptions_Category_Templates", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_DefaultKeyFileName", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_DefaultKeyFileName", typeof(Resource))]
        public string DefaultKeyFileName { get; set; } = Resource.DefaultKeyName;

        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_UseIntellisense", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_UseIntellisense", typeof(Resource))]
        public bool UseIntellisense { get; set; } = false;

        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_IntellisenseEntityTriggerCharacter", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_IntellisenseEntityTriggerCharacter", typeof(Resource))]
        public string IntellisenseEntityTriggerCharacter { get; set; } = "$";

        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_IntellisenseFieldTriggerCharacter", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_IntellisenseFieldTriggerCharacter", typeof(Resource))]
        public string IntellisenseFieldTriggerCharacter { get; set; } = "_";

        public delegate void ApplySettingsHandler(object sender, EventArgs e);
        public event PropertyChangedEventHandler PropertyChanged;

        protected override void OnApply(PageApplyEventArgs e)
        {
            if (IntellisenseEntityTriggerCharacter.Length > 1 || IntellisenseFieldTriggerCharacter.Length > 1)
            {
                MessageBox.Show(Resource.ErrorMessage_IntellisenseTriggerCharacterLength);
                e.ApplyBehavior = ApplyKind.CancelNoNavigate;
                return;
            }

            if (IntellisenseEntityTriggerCharacter.Length == 0)
                IntellisenseEntityTriggerCharacter = "$";

            if (IntellisenseFieldTriggerCharacter.Length == 0)
                IntellisenseFieldTriggerCharacter = "_";

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
                    Resource.ExtensionLoggingEnabled :
                    Resource.ExtensionLoggingDisabled);

            if (value == "XrmToolingLoggingEnabled")
                ExtensionLogger.Log(LogLevel.Info, XrmToolingLoggingEnabled ?
                    Resource.XrmToolingLoggingEnabled :
                    Resource.XrmToolingLoggingDisabled);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}