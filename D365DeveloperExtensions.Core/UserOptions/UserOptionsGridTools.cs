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
    public class UserOptionsGridTools : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private string _solutionPackagerToolPath = String.Empty;
        private string _pluginRegistrationToolPath = String.Empty;
        private string _crmSvcUtilToolPath = String.Empty;

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_RegToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_RegToolPath", typeof(Resource))]
        public string PluginRegistrationToolPath
        {
            get => _pluginRegistrationToolPath;
            set
            {
                if (_pluginRegistrationToolPath == value)
                    return;

                _pluginRegistrationToolPath = value;
                NotifyPropertyChanged();
            }
        }

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_SolutionPackToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_SolutionPackToolPath", typeof(Resource))]
        public string SolutionPackagerToolPath
        {
            get => _solutionPackagerToolPath;
            set
            {
                if (_solutionPackagerToolPath == value)
                    return;

                _solutionPackagerToolPath = value;
                NotifyPropertyChanged();
            }
        }

        //TODO: validate file path format
        [LocalizedCategory("UserOptions_Category_Tools", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_CrmSvcToolPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_CrmSvcToolPath", typeof(Resource))]
        public string CrmSvcUtilToolPath
        {
            get => _crmSvcUtilToolPath;
            set
            {
                if (_crmSvcUtilToolPath == value)
                    return;

                _crmSvcUtilToolPath = value;
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
            switch (value)
            {
                case "PluginRegistrationToolPath":
                    Logger.Log(LogLevel.Info, $"PluginRegistrationToolPath now: {PluginRegistrationToolPath}");
                    break;
                case "SolutionPackagerToolPath":
                    Logger.Log(LogLevel.Info, $"SolutionPackagerToolPath now: {SolutionPackagerToolPath}");
                    break;
                case "CrmSvcUtilToolPath":
                    Logger.Log(LogLevel.Info, $"CrmSvcUtilToolPath now: {CrmSvcUtilToolPath}");
                    break;
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}