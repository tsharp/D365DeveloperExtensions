using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public partial class NewWebResource
    {
        private readonly CrmServiceClient _client;
        private readonly DTE _dte;
        private ObservableCollection<WebResourceType> _webResourceTypes;

        public Guid NewId;
        public int NewType;
        public string NewName;
        public string NewDisplayName;
        public string NewBoundFile;
        public Guid NewSolutionId;
        public ObservableCollection<WebResourceType> WebResourceTypes
        {
            get => _webResourceTypes;
            set
            {
                _webResourceTypes = value;
                OnPropertyChanged();
            }
        }

        public NewWebResource(CrmServiceClient client, DTE dte, ObservableCollection<ComboBoxItem> projectFiles, Guid selectedSolutionId)
        {
            InitializeComponent();
            DataContext = this;
            _client = client;
            _dte = dte;

            bool result = GetSolutions(selectedSolutionId);
            if (!result)
            {
                MessageBox.Show("Error Retrieving Solutions From CRM. See the Output Window for additional details.");
                DialogResult = false;
                Close();
            }

            Files.ItemsSource = projectFiles;
            WebResourceTypes =
                CrmDeveloperExtensions2.Core.Models.WebResourceTypes.GetTypes(client.ConnectedOrgVersion.Major, false);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private async void Create_OnClick(object sender, RoutedEventArgs e)
        {
            if (!ValidateForm())
                return;

            string filePath = GetFilePath();
            if (string.IsNullOrEmpty(filePath))
                return;

            string relativePath = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
            string name = Name.Text.Trim();
            string prefix = Prefix.Text;
            int type = ((WebResourceType)Type.SelectedItem).Type;
            string displayName = DisplayName.Text.Trim();
            string description = Description.Text.Trim();

            Overlay.ShowMessage(_dte, "Creating...");

            Entity webResource =
                Crm.WebResource.CreateNewWebResourceEntity(type, prefix, name, displayName, description, filePath);

            Guid webResourceId = await Task.Run(() => Crm.WebResource.CreateWebResourceInCrm(_client, webResource));
            if (webResourceId == Guid.Empty)
            {
                Overlay.HideMessage(_dte);
                DialogResult = false;
                Close();
            }

            CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
            if (solution.SolutionId != ExtensionConstants.DefaultSolutionId)
            {
                bool addedToSolution = await Task.Run(() => Crm.Solution.AddWebResourceToSolution(_client, solution.UniqueName, webResourceId));
                if (!addedToSolution)
                {
                    Overlay.HideMessage(_dte);
                    DialogResult = false;
                    Close();
                    return;
                }
            }

            NewId = webResourceId;
            NewType = type;
            NewName = prefix + name;
            if (!string.IsNullOrEmpty(displayName))
                NewDisplayName = displayName;
            NewBoundFile = relativePath;
            NewSolutionId = solution.SolutionId;

            Overlay.HideMessage(_dte);
            DialogResult = true;
            Close();
        }

        private string GetFilePath()
        {
            ProjectItem projectItem = (ProjectItem)((ComboBoxItem)Files.SelectedItem).Tag;
            string filePath = projectItem.Properties.Item("FullPath").Value.ToString();
            if (File.Exists(filePath))
                return filePath;

            OutputLogger.WriteToOutputWindow("Missing File: " + filePath, MessageType.Error);
            MessageBox.Show("File does not exist");
            return null;
        }

        private bool ValidateForm()
        {
            if (Crm.WebResource.ValidateName(Name.Text))
                return true;

            MessageBox.Show("Web resource names may only include letters, numbers, periods, and nonconsecutive forward slash characters");
            Name.Focus();
            return false;
        }

        private void Files_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Files.SelectedItem == null)
            {
                Name.Text = String.Empty;
                DisplayName.Text = String.Empty;
                return;
            }

            string fileName = ((ComboBoxItem)Files.SelectedItem).Content.ToString();

            DisplayName.Text = FileNameToDisplayName(fileName);

            string extension = Path.GetExtension(fileName);

            if (extension.ToUpper() == ".TS")
            {
                DisplayName.Text = Path.ChangeExtension(DisplayName.Text, ".js");
                Name.Text = Path.ChangeExtension(fileName, ".js");
            }
            else
                Name.Text = fileName;

            if (string.IsNullOrEmpty(extension))
            {
                Type.SelectedValue = null;
                return;
            }

            Type.SelectedItem = WebResourceTypes.FirstOrDefault(t => t.Name == extension.Replace(".", string.Empty).ToUpper());
        }

        private string FileNameToDisplayName(string fileName)
        {
            if (fileName.Count(s => s == '/') != 1) //nested in folder
                return fileName;

            fileName = fileName.Replace("/", String.Empty);

            if (fileName.StartsWith(Prefix.Text, StringComparison.InvariantCultureIgnoreCase))
                fileName = fileName.Substring(Prefix.Text.Length, fileName.Length - Prefix.Text.Length);

            return fileName;
        }

        private void Solutions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Solutions.SelectedItem != null)
            {
                CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
                Prefix.Text = solution.Prefix + "_";
            }
            else
                Prefix.Text = "new_";
        }

        private bool GetSolutions(Guid selectedSolutionId)
        {
            Overlay.ShowMessage(_dte, "Getting solutions from CRM...");

            EntityCollection results = Crm.Solution.RetrieveSolutionsFromCrm(_client, false);

            List<CrmSolution> solutions = ModelBuilder.CreateCrmSolutionView(results);

            if (selectedSolutionId != Guid.Empty)
            {
                var sel = solutions.FindIndex(s => s.SolutionId == selectedSolutionId);
                if (sel != -1)
                    Solutions.SelectedIndex = sel;
            }
            else
                Solutions.SelectedIndex = 0;

            Solutions.ItemsSource = solutions;

            Overlay.HideMessage(_dte);

            return true;
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}