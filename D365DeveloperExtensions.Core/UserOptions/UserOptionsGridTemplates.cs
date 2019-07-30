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
    public class UserOptionsGridTemplates : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private string _customTemplatesPath = string.Empty;
        private string _defaultKeyFileName = Resource.DefaultKeyName;

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
        [LocalizedCategory("UserOptions_Category_Templates", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_TemplatesPath", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_TemplatesPath", typeof(Resource))]
        public string CustomTemplatesPath
        {
            get => _customTemplatesPath;
            set
            {
                if (_customTemplatesPath == value)
                    return;

                _customTemplatesPath = value;
                NotifyPropertyChanged();
            }
        }

        [LocalizedCategory("UserOptions_Category_Templates", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_DefaultKeyFileName", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_DefaultKeyFileName", typeof(Resource))]
        public string DefaultKeyFileName
        {
            get => _defaultKeyFileName;
            set
            {
                if (_defaultKeyFileName == value)
                    return;

                _defaultKeyFileName = value;
                NotifyPropertyChanged();
            }
        }

        public delegate void ApplySettingsHandler(object sender, EventArgs e);
        public event ApplySettingsHandler Applied;
        public event PropertyChangedEventHandler PropertyChanged;

        protected override void OnApply(PageApplyEventArgs e)
        {
            if (!FileSystem.IsValidFolder(CustomTemplatesPath))
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
                case "CustomTemplatesPath":
                    SharedGlobals.SetGlobal("CustomTemplatesPath", CustomTemplatesPath);
                    Logger.Log(LogLevel.Info, $"CustomTemplatesPath now: {CustomTemplatesPath}");
                    break;
                case "DefaultKeyFileName":
                    SharedGlobals.SetGlobal("DefaultKeyFileName", DefaultKeyFileName);
                    Logger.Log(LogLevel.Info, $"DefaultKeyFileName now: {DefaultKeyFileName}");
                    break;
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}