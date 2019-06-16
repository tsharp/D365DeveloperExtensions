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
    public class UserOptionsGridIntellisense : DialogPage, INotifyPropertyChanged
#pragma warning restore CS3021 // Type or member does not need a CLSCompliant attribute because the assembly does not have a CLSCompliant attribute
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private bool _useIntellisense;
        private string _intellisenseEntityTriggerCharacter = "$";
        private string _intellisenseFieldTriggerCharacter = "_";

        //Need to add entry to D365DeveloperExtensions - D365DeveloperExtensionsPackage.cs [ProvideOptionPage]
        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_UseIntellisense", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_UseIntellisense", typeof(Resource))]
        public bool UseIntellisense
        {
            get => _useIntellisense;
            set
            {
                if (_useIntellisense == value)
                    return;

                _useIntellisense = value;
                NotifyPropertyChanged();
            }
        }

        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_IntellisenseEntityTriggerCharacter", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_IntellisenseEntityTriggerCharacter", typeof(Resource))]
        public string IntellisenseEntityTriggerCharacter
        {
            get => _intellisenseEntityTriggerCharacter;
            set
            {
                if (_intellisenseEntityTriggerCharacter == value)
                    return;

                _intellisenseEntityTriggerCharacter = value;
                NotifyPropertyChanged();
            }
        }

        [LocalizedCategory("UserOptions_Category_Intellisense", typeof(Resource))]
        [LocalizedDisplayName("UserOptions_DisplayName_IntellisenseFieldTriggerCharacter", typeof(Resource))]
        [LocalizedDescription("UserOptions_Description_IntellisenseFieldTriggerCharacter", typeof(Resource))]
        public string IntellisenseFieldTriggerCharacter
        {
            get => _intellisenseFieldTriggerCharacter;
            set
            {
                if (_intellisenseFieldTriggerCharacter == value)
                    return;

                _intellisenseFieldTriggerCharacter = value;
                NotifyPropertyChanged();
            }
        }

        public delegate void ApplySettingsHandler(object sender, EventArgs e);
        public event ApplySettingsHandler Applied;
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

        private void OnApplied(PageApplyEventArgs e)
        {
            Applied?.Invoke(this, e);
        }

        private void NotifyPropertyChanged([CallerMemberName] string value = null)
        {
            switch (value)
            {
                case "UseIntellisense":
                    SharedGlobals.SetGlobal("UseIntellisense", UseIntellisense);
                    Logger.Log(LogLevel.Info, $"UseIntellisense now: {UseIntellisense}");
                    break;
                case "IntellisenseEntityTriggerCharacter":
                    SharedGlobals.SetGlobal("IntellisenseEntityTriggerCharacter", IntellisenseEntityTriggerCharacter);
                    Logger.Log(LogLevel.Info, $"IntellisenseEntityTriggerCharacter now: {IntellisenseEntityTriggerCharacter}");
                    break;
                case "IntellisenseFieldTriggerCharacter":
                    SharedGlobals.SetGlobal("IntellisenseFieldTriggerCharacter", IntellisenseFieldTriggerCharacter);
                    Logger.Log(LogLevel.Info, $"IntellisenseFieldTriggerCharacter now: {IntellisenseFieldTriggerCharacter}");
                    break;
            }

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(value));
        }
    }
}