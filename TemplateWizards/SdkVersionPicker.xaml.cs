using System.Collections.Generic;
using System.Windows.Controls;
using Common.Models;
using Microsoft.VisualStudio.PlatformUI;
using NuGetRetriever;
using TemplateWizards.Resources;

namespace TemplateWizards
{
    public partial class SdkVersionPicker : DialogWindow
    {
        public SdkVersionPicker()
        {
            //Resource.Culture = new System.Globalization.CultureInfo("it-IT");
            InitializeComponent();

            GetCrmPackages();
        }

        private void GetCrmPackages()
        {
            List<NuGetPackage> packages = PackageLister.GetPackagesbyId("Microsoft.CrmSdk.CoreAssemblies");

            foreach (NuGetPackage nuGetPackage in packages)
            {
                ListViewItem item = new ListViewItem
                {
                    Content = nuGetPackage.VersionText,
                    Tag = nuGetPackage
                };

                SdkVersions.Items.Add(item);
            }
        }
    }
}
