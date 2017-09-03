using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class SolutionsButton
    {
        public SolutionsButton()
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
            "IsConnected", typeof(bool), typeof(SolutionsButton),
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
            SolutionsButton solutionsButton = d as SolutionsButton;
            solutionsButton?.OnIsConnectedChange(e);
        }

        private void OnIsConnectedChange(DependencyPropertyChangedEventArgs e)
        {
            IsConnected = (bool)e.NewValue;
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            if (!(SharedGlobals.GetGlobal("CrmService", dte) is CrmServiceClient client))
            {
                OutputLogger.WriteToOutputWindow("Not connected to CRM/365.", MessageType.Error);
                return;
            }

            WebBrowser.OpenCrmPage(dte, client,
                "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void Solutions_OnLoaded(object sender, RoutedEventArgs e)
        {
            DTE dte = (DTE)Package.GetGlobalService(typeof(DTE));
            if (dte == null)
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