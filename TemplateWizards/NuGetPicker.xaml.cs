using CrmDeveloperExtensions2.Core.Models;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

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

            _packageName = packageName;
            _packageVersions = packageVersions;

            GetPackage(packageName, packageVersions);
        }

        private void GetPackage(string packageName, List<NuGetPackage> packageVersions)
        {
            Versions.Items.Clear();

            if (LimitVersions.IsChecked != null && LimitVersions.IsChecked.Value)
                packageVersions = FilterLatestVersions(packageVersions);

            VersionsGrid.Columns[0].Header = packageName;

            foreach (NuGetPackage package in packageVersions)
            {
                var item = CreateItem(package);

                Versions.Items.Add(item);
            }

            Versions.SelectedIndex = 0;
        }

        private static ListViewItem CreateItem(NuGetPackage package)
        {
            ListViewItem item = new ListViewItem
            {
                Content = package.VersionText,
                Tag = package
            };

            return item;
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
            ListView versions = (ListView)sender;
            ListBoxItem item = versions.SelectedItem as ListViewItem;
            if (item == null)
                return;

            SelectedPackage = item.Tag as NuGetPackage;
        }

        private static List<NuGetPackage> FilterLatestVersions(List<NuGetPackage> packageVersions)
        {
            List<NuGetPackage> filteredNuGetPackages = new List<NuGetPackage>();

            Version firstVersion = packageVersions[0].Version;
            var currentMajor = firstVersion.Major;
            var currentMinor = firstVersion.Minor;
            var currentPackage = packageVersions[0];

            for (int i = 0; i < packageVersions.Count; i++)
            {
                if (i == packageVersions.Count - 1)
                {
                    filteredNuGetPackages.Add(currentPackage);
                    continue;
                }

                Version ver = packageVersions[i].Version;

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
    }
}