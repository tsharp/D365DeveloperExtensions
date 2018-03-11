using D365DeveloperExtensions.Core.Localization;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;

namespace D365DeveloperExtensions.Core.UserOptions
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
#pragma warning disable CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    public class UserOptionsGridLogging : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private bool _extensionLoggingEnabled;
        private bool _xrmToolingLoggingEnabled;
        private string _extensionLogFilePath = String.Empty;
        private string _xrmToolingLogFilePath = String.Empty;

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
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

        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_ExtensionLoggingPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_ExtensionLoggingPath", typeof(Resource))]
        public string ExtensionLogFilePath
        {
            get => _extensionLogFilePath;
            set
            {
                if (_extensionLogFilePath == value)
                    return;

                _extensionLogFilePath = value;
                NotifyPropertyChanged();
            }
        }

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

        [LocalizedCategory("UserOptions_Category_Logging", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_XrmToolingLoggingPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_XrmToolingLoggingPath", typeof(Resource))]
        public string XrmToolingLogFilePath
        {
            get => _xrmToolingLogFilePath;
            set
            {
                if (_xrmToolingLogFilePath == value)
                    return;

                _xrmToolingLogFilePath = value;
                NotifyPropertyChanged();
            }
        }

        public delegate void ApplySettingsHandler(object sender, EventArgs e);
        public event ApplySettingsHandler Applied;
        public event PropertyChangedEventHandler PropertyChanged;

        protected override void OnApply(PageApplyEventArgs e)
        {
            if (!FileSystem.IsValidFolder(ExtensionLogFilePath) || !FileSystem.IsValidFolder(XrmToolingLogFilePath))
            {
                MessageBox.Show(Resource.Error_InvalidFolderPath);
                e.ApplyBehavior = ApplyKind.CancelNoNavigate;
                return;
            }

            OnApplied(e);
            base.OnApply(e);
        }

        private void OnApplied(PageApplyEventArgs e)
        {
            Applied?.Invoke(this, e);
        }

        private void NotifyPropertyChanged([CallerMemberName] string value = null)
        {
            switch (value)
            {
                case "ExtensionLoggingEnabled":
                    Logger.Log(LogLevel.Info, $"ExtensionLoggingEnabled now: {ExtensionLoggingEnabled}");
                    break;
                case "ExtensionLogFilePath":
                    Logger.Log(LogLevel.Info, $"ExtensionLogFilePath now: {ExtensionLogFilePath}");
                    break;
                case "XrmToolingLoggingEnabled":
                    Logger.Log(LogLevel.Info, $"XrmToolingLoggingEnabled now: {XrmToolingLoggingEnabled}");
                    break;
                case "XrmToolingLogFilePath":
                    Logger.Log(LogLevel.Info, $"XrmToolingLogFilePath now: {XrmToolingLogFilePath}");
                    break;
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}