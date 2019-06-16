using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace D365DeveloperExtensions.Core.Controls
{
    public partial class CustomizationButton : INotifyPropertyChanged
    {
        public CustomizationButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static readonly DependencyProperty IsConnectedProperty = DependencyProperty.Register(
            "IsConnected", typeof(bool), typeof(CustomizationButton),
            new PropertyMetadata(default(bool), OnIsConnectedChange));

        public bool IsConnected
        {
            get => (bool)GetValue(IsConnectedProperty);

            set
            {
                SetValue(IsConnectedProperty, value);
                OnPropertyChanged();
            }
        }

        private static void OnIsConnectedChange(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CustomizationButton customizationButton = d as CustomizationButton;
            customizationButton?.OnIsConnectedChange(e);
        }

        private void OnIsConnectedChange(DependencyPropertyChangedEventArgs e)
        {
            IsConnected = (bool)e.NewValue;
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            if (!(SharedGlobals.GetGlobal("CrmService", dte) is CrmServiceClient client))
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_NotConnectedOrg, MessageType.Error);
                return;
            }

            WebBrowser.OpenCrmPage(client, $"tools/solution/edit.aspx?id=%7b{ExtensionConstants.DefaultSolutionId}%7d");
        }

        private void Customizations_OnLoaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            if (SharedGlobals.GetGlobal("CrmService", dte) is CrmServiceClient client)
                IsConnected = client.ConnectedOrgUniqueName != null;
            else
            {
                IsConnected = false;
                SetBinding(IsEnabledProperty, "IsConnected");
            }
        }
    }
}