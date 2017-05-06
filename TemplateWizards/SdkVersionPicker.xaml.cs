using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Common.Models;
using Microsoft.VisualStudio.PlatformUI;
using NuGetRetriever;
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

            GetPackage("Microsoft.CrmSdk.CoreAssemblies");
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

        private void OK_OnClick(object sender, RoutedEventArgs e)
        {
            Package = (PackageValue)NuGetProcessor.GetNextPackage(Package, GetWorkflow, GetClient);

            switch (Package)
            {
                case PackageValue.Workflow:
                    GetPackage("Microsoft.CrmSdk.Workflow");
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
