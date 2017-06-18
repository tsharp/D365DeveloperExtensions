using CrmDeveloperExtensions2.Core.Models;
using Microsoft.VisualStudio.PlatformUI;
using NuGetRetriever;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace TemplateWizards
{
    public partial class NuGetPicker : DialogWindow
    {
        public NuGetPackage SelectedPackage { get; set; }

        public NuGetPicker(string packageId)
        {
            //Resource.Culture = new System.Globalization.CultureInfo("it-IT");
            InitializeComponent();


            GetPackage(packageId);
        }

        private void GetPackage(string nuGetPackage)
        {
            List<NuGetPackage> versions = PackageLister.GetPackagesbyId(nuGetPackage);

            VersionsGrid.Columns[0].Header = nuGetPackage;

            foreach (NuGetPackage package in versions)
            {
                ListViewItem item = new ListViewItem
                {
                    Content = package.VersionText,
                    Tag = package
                };

                Versions.Items.Add(item);
            }
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
    }
}