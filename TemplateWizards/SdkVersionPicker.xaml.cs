using CrmDeveloperExtensions2.Core.Models;
using Microsoft.VisualStudio.PlatformUI;
using NuGetRetriever;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using TemplateWizards.Enums;
using TemplateWizards.Resources;

namespace TemplateWizards
{
    public partial class SdkVersionPicker : DialogWindow
    {
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
            List<NuGetPackage> versions = PackageLister.GetPackagesbyId(nuGetPackage);

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

            //TODO: need to ensure a default value is picked after loading the package list
            //string selectedVersion = ((ListViewItem)SdkVersions.SelectedItem).Content.ToString();
            //SetSelectedVersion(selectedVersion);
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
            string selectedVersion = ((ListViewItem)sdkVersions.SelectedItem).Content.ToString();
            SetSelectedVersion(selectedVersion);
        }
    }
}