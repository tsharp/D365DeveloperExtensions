using CrmDeveloperExtensions2.Core.Models;
using NuGetRetriever;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using TemplateWizards.Enums;
using TemplateWizards.Resources;

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
        private bool GetWorkflow { get; set; }
        public bool GetClient { get; set; }
        public PackageValue Package = PackageValue.Core;

        public SdkVersionPicker(bool getWorkflow, bool getClient)
        {
            //Resource.Culture = new System.Globalization.CultureInfo("it-IT");
            InitializeComponent();

            GetWorkflow = getWorkflow;
            GetClient = getClient;

            GetPackage(Resource.SdkAssemblyCore);
        }

        private void GetPackage(string nuGetPackage)
        {
            SdkVersions.Items.Clear();

            Title = $"Choose Version:  {nuGetPackage}";

            List<NuGetPackage> versions = PackageLister.GetPackagesbyId(nuGetPackage);

            _packageVersions = versions;
            _currentPackage = nuGetPackage;

            if (LimitVersions.IsChecked != null && LimitVersions.IsChecked.Value)
                versions = FilterLatestVersions(versions);

            SdkVersionsGrid.Columns[0].Header = nuGetPackage;

            foreach (NuGetPackage package in versions)
            {
                ListViewItem item = new ListViewItem
                {
                    Content = package.VersionText,
                    Tag = package
                };

                SdkVersions.Items.Add(item);
            }

            SdkVersions.SelectedIndex = 0;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CloseDialog(false);
        }

        private void SetSelectedVersion(string selectedVersion)
        {
            switch (Package)
            {
                case PackageValue.Core:
                    CoreVersion = selectedVersion;
                    break;
                case PackageValue.Workflow:
                    WorkflowVersion = selectedVersion;
                    break;
                case PackageValue.Client:
                    ClientVersion = selectedVersion;
                    break;
            }
        }

        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            Package = (PackageValue)NuGetProcessor.GetNextPackage(Package, GetWorkflow, GetClient);

            switch (Package)
            {
                case PackageValue.Workflow:
                    GetPackage(Resource.SdkAssemblyWorkflow);
                    break;
                case PackageValue.Client:
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
            ListView sdkVersions = (ListView)sender;
            ListBoxItem item = sdkVersions.SelectedItem as ListViewItem;
            if (item == null)
                return;

            string selectedVersion = ((ListViewItem)sdkVersions.SelectedItem).Content.ToString();
            SetSelectedVersion(selectedVersion);
        }

        private static List<NuGetPackage> FilterLatestVersions(List<NuGetPackage> versions)
        {
            List<NuGetPackage> filteredVersions = new List<NuGetPackage>();

            Version firstVersion = versions[0].Version;
            var currentMajor = firstVersion.Major;
            var currentMinor = firstVersion.Minor;
            var currentPackage = versions[0];

            for (int i = 0; i < versions.Count; i++)
            {
                if (i == versions.Count - 1)
                {
                    filteredVersions.Add(currentPackage);
                    continue;
                }

                Version ver = versions[i].Version;

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
    }
}