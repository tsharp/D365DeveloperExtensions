using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NLog;
using Logger = NLog.Logger;

namespace CrmDeveloperExtensions.Core
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
#pragma warning disable CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    public class UserOptionsGrid : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private bool _loggingEnabled;

        public delegate void ApplySettingsHandler(object sender, EventArgs e); //Applied on save of form -may not be required
        public event PropertyChangedEventHandler PropertyChanged;

        [Category("Logging Options")]
        [DisplayName("Enable detailed extension logging?")]
        [Description("Detailed extension logging will log a bunch of stuff")]
        public bool LoggingEnabled
        {
            get => _loggingEnabled;
            set
            {
                if (_loggingEnabled == value)
                    return;

                _loggingEnabled = value;
                NotifyPropertyChanged();
            }
        }

        [Category("Logging Options")]
        [DisplayName("Log file path")]
        [Description("Path to log file storage")]
        public string LogFilePath { get; set; } = String.Empty;

        private void NotifyPropertyChanged([CallerMemberName] string value = null)
        {
            if (value == "LoggingEnabled")
                ExtensionLogger.Log(LogLevel.Info, LoggingEnabled ?
                    Resources.Resource.FileLoggingEnabled :
                    Resources.Resource.FileLoggingDisabled);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }

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
    }
}
