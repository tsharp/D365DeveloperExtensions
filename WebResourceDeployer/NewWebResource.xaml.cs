using EnvDTE;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.VisualStudio.PlatformUI;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public partial class NewWebResource
    {
        private readonly CrmServiceClient _client;

        public Guid NewId;
        public int NewType;
        public string NewName;
        public string NewDisplayName;
        public string NewBoundFile;
        public Guid NewSolutionId;

        public NewWebResource(CrmServiceClient client, ObservableCollection<ComboBoxItem> projectFiles, Guid selectedSolutionId)
        {
            InitializeComponent();
            _client = client;

            bool result = GetSolutions(selectedSolutionId);
            if (!result)
            {
                MessageBox.Show("Error Retrieving Solutions From CRM. See the Output Window for additional details.");
                DialogResult = false;
                Close();
            }

            Files.ItemsSource = projectFiles;
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
            int type = Convert.ToInt32(((ComboBoxItem)Type.SelectedItem).Tag.ToString());
            string displayName = DisplayName.Text.Trim();
            string description = Description.Text.Trim();

            ShowMessage("Creating...");

            Entity webResource =
                Crm.WebResource.CreateNewWebResourceEntity(type, prefix, name, displayName, description, filePath);

            Guid webResourceId = await Task.Run(() => Crm.WebResource.CreateWebResourceInCrm(_client, webResource));
            if (webResourceId == Guid.Empty)
            {
                HideMessage();
                DialogResult = false;
                Close();
            }

            CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
            if (solution.SolutionId != ExtensionConstants.DefaultSolutionId)
            {
                bool addedToSolution = await Task.Run(() => Crm.Solution.AddWebResourceToSolution(_client, solution.UniqueName, webResourceId));
                if (!addedToSolution)
                {
                    HideMessage();
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

            HideMessage();
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
            if (Files.SelectedItem == null)
            {
                MessageBox.Show("Select a file");
                Files.Focus();
                return false;
            }

            if (string.IsNullOrEmpty(Name.Text))
            {
                MessageBox.Show("Name is required");
                Name.Focus();
                return false;
            }

            if (!Crm.WebResource.ValidateName(Name.Text))
            {
                MessageBox.Show("Web resource names may only include letters, numbers, periods, and nonconsecutive forward slash characters");
                Name.Focus();
                return false;
            }

            if (Type.SelectedItem == null)
            {
                MessageBox.Show("Select a web resource type");
                Type.Focus();
                return false;
            }

            return true;
        }

        private void Type_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //TypeLabel.Foreground = Type.SelectedItem != null ? Brushes.Black : Brushes.Red;
        }

        private void Name_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            //NameLabel.Foreground = string.IsNullOrEmpty(Name.Text) ? Brushes.Red : Brushes.Black;
        }

        private void Files_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Files.SelectedItem == null)
            {
                //FilesLabel.Foreground = Brushes.Red;
                Name.Text = String.Empty;
                DisplayName.Text = String.Empty;
                return;
            }

            //FilesLabel.Foreground = Brushes.Black;
            string fileName = ((ComboBoxItem)Files.SelectedItem).Content.ToString();

            DisplayName.Text = FileNameToDisplayName(fileName);

            string extension = Path.GetExtension(fileName);

            if (extension.ToUpper() != ".TS")
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

            Type.SelectedValue = ExtensionToType(extension);
        }

        private string ExtensionToType(string extension)
        {
            switch (extension.ToUpper())
            {
                case ".CSS":
                    return "Style Sheet (CSS)";
                case ".JS":
                case ".TS":
                    return "Script (JScript)";
                case ".XML":
                    return "Data (XML)";
                case ".PNG":
                    return "PNG format";
                case ".JPG":
                    return "JPG format";
                case ".GIF":
                    return "GIF format";
                case ".XAP":
                    return "Silverlight (XAP)";
                case ".XSL":
                case ".XSLT":
                    return "Style Sheet (XSL)";
                case ".ICO":
                    return "ICO format";
                default:
                    return "Webpage (HTML)";
            }
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
            bool isSolutionSelected = Solutions.SelectedItem != null;
            if (Solutions.SelectedItem != null)
            {
                CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
                //SolutionsLabel.Foreground = Brushes.Black;
                Prefix.Text = solution.Prefix + "_";
            }
            else
            {
                //SolutionsLabel.Foreground = Brushes.Red;
                Prefix.Text = "new_";
            }

            Files.IsEnabled = isSolutionSelected;
            Name.IsEnabled = isSolutionSelected;
            DisplayName.IsEnabled = isSolutionSelected;
            Type.IsEnabled = isSolutionSelected;
            Description.IsEnabled = isSolutionSelected;

            Name.Text = String.Empty;
            DisplayName.Text = String.Empty;
            Type.SelectedIndex = -1;
            Files.SelectedIndex = -1;
        }

        private bool GetSolutions(Guid selectedSolutionId)
        {
            ShowMessage("Getting solutions from CRM...");

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

            HideMessage();

            return true;
        }

        private void ShowMessage(string message)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        //LockMessage.Content = message;
                        //LockOverlay.Visibility = Visibility.Visible;
                    }
                ));
        }

        private void HideMessage()
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        //LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}