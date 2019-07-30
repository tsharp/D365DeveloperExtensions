using D365DeveloperExtensions.Core.ExtensionMethods;
using D365DeveloperExtensions.Core.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using TemplateWizards.Resources;

namespace TemplateWizards
{
    public partial class NuGetPicker
    {
        private readonly string _packageName;
        private readonly List<NuGetPackage> _packageVersions;
        public NuGetPackage SelectedPackage { get; set; }

        public NuGetPicker(string packageName, List<NuGetPackage> packageVersions)
        {
            InitializeComponent();
            Owner = Application.Current.MainWindow;
            LicenseLink.Inlines.Add(Resource.NuGetPicker_LicenseInfoText_TextBlock_Text);

            _packageName = packageName;
            _packageVersions = packageVersions;

            Title += $"{Resource.Version_Window_Title}: {packageName}";

            GetPackage(packageName, packageVersions);
        }

        private void GetPackage(string packageName, List<NuGetPackage> packageVersions)
        {
            Versions.Items.Clear();

            if (LimitVersions.ReturnValue())
                packageVersions = FilterLatestVersions(packageVersions);

            VersionsGrid.Columns[0].Header = packageName;

            foreach (var package in packageVersions)
            {
                Versions.Items.Add(package);
            }

            Versions.SelectedIndex = 0;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CloseDialog(false);
        }

        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            CloseDialog(true);
        }

        private void CloseDialog(bool result)
        {
            DialogResult = result;
            Close();
        }

        private void Versions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sdkVersions = (ListView)sender;
            if (!(sdkVersions.SelectedItem is NuGetPackage package))
                return;

            //TODO: set selected version

            if (!string.IsNullOrEmpty(package.LicenseUrl))
            {
                LicensePanel.Visibility = Visibility.Visible;
                LicenseLink.NavigateUri = new Uri(package.LicenseUrl);
            }
            else
                LicensePanel.Visibility = Visibility.Hidden;
        }

        private static List<NuGetPackage> FilterLatestVersions(List<NuGetPackage> packageVersions)
        {
            var filteredNuGetPackages = new List<NuGetPackage>();

            var firstVersion = packageVersions[0].Version;
            var currentMajor = firstVersion.Major;
            var currentMinor = firstVersion.Minor;
            var currentPackage = packageVersions[0];

            for (var i = 0; i < packageVersions.Count; i++)
            {
                if (i == packageVersions.Count - 1)
                {
                    filteredNuGetPackages.Add(currentPackage);
                    continue;
                }

                var ver = packageVersions[i].Version;

                if (ver.Major < currentMajor)
                {
                    currentMajor = ver.Major;
                    currentMinor = ver.Minor;
                    filteredNuGetPackages.Add(currentPackage);
                    currentPackage = packageVersions[i];
                    continue;
                }

                if (ver.Minor < currentMinor)
                {
                    currentMinor = ver.Minor;
                    filteredNuGetPackages.Add(currentPackage);
                    currentPackage = packageVersions[i];
                }
            }

            return filteredNuGetPackages;
        }

        private void LimitVersions_Checked(object sender, RoutedEventArgs e)
        {
            if (_packageVersions != null)
                GetPackage(_packageName, _packageVersions);
        }

        private void LicenseLink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}