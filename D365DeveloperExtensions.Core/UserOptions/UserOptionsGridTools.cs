using D365DeveloperExtensions.Core.Localization;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;

namespace D365DeveloperExtensions.Core.UserOptions
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
#pragma warning disable CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    public class UserOptionsGridTools : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private string _solutionPackagerToolPath = string.Empty;
        private string _pluginRegistrationToolPath = string.Empty;
        private string _crmSvcUtilToolPath = string.Empty;

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
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
            if (!DoesToolExist(PluginRegistrationToolPath, "PluginRegistration.exe"))
            {
                MessageBox.Show(Resource.Error_CouldNotLocatePluginRegistrationTool);
                e.ApplyBehavior = ApplyKind.CancelNoNavigate;
                return;
            }

            if (!DoesToolExist(SolutionPackagerToolPath, "SolutionPackager.exe"))
            {
                MessageBox.Show(Resource.Error_CouldNotLocateSolutionPackager);
                e.ApplyBehavior = ApplyKind.CancelNoNavigate;
                return;
            }

            if (!DoesToolExist(CrmSvcUtilToolPath, "CrmSvcUtil.exe"))
            {
                MessageBox.Show(Resource.Error_CouldNotLocateCrmSvcUtil);
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
                case "PluginRegistrationToolPath":
                    SharedGlobals.SetGlobal("PluginRegistrationToolPath", PluginRegistrationToolPath);
                    Logger.Log(LogLevel.Info, $"PluginRegistrationToolPath now: {PluginRegistrationToolPath}");
                    break;
                case "SolutionPackagerToolPath":
                    SharedGlobals.SetGlobal("SolutionPackagerToolPath", SolutionPackagerToolPath);
                    Logger.Log(LogLevel.Info, $"SolutionPackagerToolPath now: {SolutionPackagerToolPath}");
                    break;
                case "CrmSvcUtilToolPath":
                    SharedGlobals.SetGlobal("CrmSvcUtilToolPath", CrmSvcUtilToolPath);
                    Logger.Log(LogLevel.Info, $"CrmSvcUtilToolPath now: {CrmSvcUtilToolPath}");
                    break;
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }

        private static bool DoesToolExist(string folder, string tool)
        {
            if (string.IsNullOrEmpty(folder))
                return true;

            try
            {
                FileInfo file = new FileInfo(Path.Combine(folder, tool));

                return file.Exists;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}