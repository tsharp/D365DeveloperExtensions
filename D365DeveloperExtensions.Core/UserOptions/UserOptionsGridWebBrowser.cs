using D365DeveloperExtensions.Core.Localization;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace D365DeveloperExtensions.Core.UserOptions
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
#pragma warning disable CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    public class UserOptionsGridWebBrowser : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private bool _useInternalBrowser;

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
        [LocalizedCategory("UserOptions_Category_WebBrowser", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_WebBrowser", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_WebBrowser", typeof(Resource))]
        public bool UseInternalBrowser
        {
            get => _useInternalBrowser;
            set
            {
                if (_useInternalBrowser == value)
                    return;

                _useInternalBrowser = value;
                NotifyPropertyChanged();
            }
        }

        public delegate void ApplySettingsHandler(object sender, EventArgs e);
        public event ApplySettingsHandler Applied;
        public event PropertyChangedEventHandler PropertyChanged;

        protected override void OnApply(PageApplyEventArgs e)
        {
            OnApplied(e);
            base.OnApply(e);
        }

        private void OnApplied(PageApplyEventArgs e)
        {
            Applied?.Invoke(this, e);
        }

        private void NotifyPropertyChanged([CallerMemberName] string value = null)
        {
            if (value == "UseInternalBrowser")
            {
                SharedGlobals.SetGlobal("UseInternalBrowser", UseInternalBrowser);
                Logger.Log(LogLevel.Info, $"UseInternalBrowser now: {UseInternalBrowser}");
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}