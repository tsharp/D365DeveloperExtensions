using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.ExtensionMethods;
using D365DeveloperExtensions.Core.Models;
using NuGetRetriever;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using TemplateWizards.Resources;
using Process = System.Diagnostics.Process;

namespace TemplateWizards
{
    public partial class SdkVersionPicker
    {
        private List<NuGetPackage> _packageVersions;
        private string _currentPackage;

        public string CoreVersion { get; set; }
        public string WorkflowVersion { get; set; }
        public string ClientPackage { get; set; }
        public string ClientVersion { get; set; }
        private bool GetWorkflow { get; }
        public bool GetClient { get; set; }
        public TemplatePackageType TemplatePackage = TemplatePackageType.Core;

        public SdkVersionPicker(bool getWorkflow, bool getClient)
        {
            InitializeComponent();
            Owner = Application.Current.MainWindow;
            LicenseLink.Inlines.Add(Resource.SdkVersionPicker_LicenseInfoText_TextBlock_Text);

            GetWorkflow = getWorkflow;
            GetClient = getClient;

            GetPackage(Resource.SdkAssemblyCore);
        }

        private void GetPackage(string nuGetPackage)
        {
            SdkVersions.Items.Clear();
            Title = $"{Resource.Version_Window_Title}:  {nuGetPackage}";

            var versions = PackageLister.GetPackagesById(nuGetPackage);

            _packageVersions = versions;
            _currentPackage = nuGetPackage;

            if (LimitVersions.ReturnValue())
                versions = FilterLatestVersions(versions);

            SdkVersionsGrid.Columns[0].Header = nuGetPackage;

            foreach (var package in versions)
            {
                SdkVersions.Items.Add(package);
            }

            SdkVersions.SelectedIndex = 0;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CloseDialog(false);
        }

        private void SetSelectedVersion(string selectedVersion)
        {
            switch (TemplatePackage)
            {
                case TemplatePackageType.Core:
                    CoreVersion = selectedVersion;
                    break;
                case TemplatePackageType.Workflow:
                    WorkflowVersion = selectedVersion;
                    break;
                case TemplatePackageType.Client:
                    ClientVersion = selectedVersion;
                    break;
            }
        }

        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            TemplatePackage = (TemplatePackageType)NuGetProcessor.GetNextPackage(TemplatePackage, GetWorkflow, GetClient);

            switch (TemplatePackage)
            {
                case TemplatePackageType.Workflow:
                    GetPackage(Resource.SdkAssemblyWorkflow);
                    break;
                case TemplatePackageType.Client:
                    ClientPackage = NuGetProcessor.DetermineClientType(CoreVersion);
                    GetPackage(ClientPackage);
                    break;
                default:
                    CloseDialog(true);
                    break;
            }
        }

        private void CloseDialog(bool result)
        {
            DialogResult = result;
            Close();
        }

        private void SdkVersions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sdkVersions = (ListView)sender;
            if (!(sdkVersions.SelectedItem is NuGetPackage package))
                return;

            var selectedVersion = package.VersionText;
            SetSelectedVersion(selectedVersion);

            if (!string.IsNullOrEmpty(package.LicenseUrl))
            {
                LicensePanel.Visibility = Visibility.Visible;
                LicenseLink.NavigateUri = new Uri(package.LicenseUrl);
            }
            else
                LicensePanel.Visibility = Visibility.Hidden;
        }

        private static List<NuGetPackage> FilterLatestVersions(List<NuGetPackage> versions)
        {
            var filteredVersions = new List<NuGetPackage>();

            var firstVersion = versions[0].Version;
            var currentMajor = firstVersion.Major;
            var currentMinor = firstVersion.Minor;
            var currentPackage = versions[0];

            for (var i = 0; i < versions.Count; i++)
            {
                if (i == versions.Count - 1)
                {
                    filteredVersions.Add(currentPackage);
                    continue;
                }

                var ver = versions[i].Version;

                if (ver.Major < currentMajor)
                {
                    currentMajor = ver.Major;
                    currentMinor = ver.Minor;
                    filteredVersions.Add(currentPackage);
                    currentPackage = versions[i];
                    continue;
                }

                if (ver.Minor < currentMinor)
                {
                    currentMinor = ver.Minor;
                    filteredVersions.Add(currentPackage);
                    currentPackage = versions[i];
                }
            }

            return filteredVersions;
        }

        private void LimitVersions_Checked(object sender, RoutedEventArgs e)
        {
            if (_packageVersions != null)
                GetPackage(_currentPackage);
        }

        private void LicenseLink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}