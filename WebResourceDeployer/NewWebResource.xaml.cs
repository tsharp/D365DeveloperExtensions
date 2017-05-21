using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using EnvDTE;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.Internal.VisualStudio.PlatformUI;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    /// <summary>
    /// Interaction logic for NewWebResource.xaml
    /// </summary>
    public partial class NewWebResource : DialogWindow
    {
        private readonly CrmServiceClient _client;

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

        private void Create_OnClick(object sender, RoutedEventArgs e)
        {


            DialogResult = true;
            Close();
        }

        private void Type_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TypeLabel.Foreground = Type.SelectedItem != null ? Brushes.Black : Brushes.Red;
        }

        private void Name_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            NameLabel.Foreground = string.IsNullOrEmpty(Name.Text) ? Brushes.Red : Brushes.Black;
        }

        private void Files_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Solutions_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Solutions.SelectedItem != null)
            {
                CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
                SolutionsLabel.Foreground = Brushes.Black;
                Prefix.Text = solution.Prefix + "_";
                Files.IsEnabled = true;
                Name.IsEnabled = true;
                DisplayName.IsEnabled = true;
                Type.IsEnabled = true;
                Description.IsEnabled = true;
            }
            else
            {
                SolutionsLabel.Foreground = Brushes.Red;
                Prefix.Text = "new_";
                Files.IsEnabled = false;
                Name.IsEnabled = false;
                DisplayName.IsEnabled = false;
                Type.IsEnabled = false;
                Description.IsEnabled = false;
            }

            Name.Text = null;
            DisplayName.Text = null;
            Type.SelectedIndex = -1;
            Files.SelectedIndex = -1;
        }

        private bool GetSolutions(Guid selectedSolutionId)
        {
            ShowMessage("Getting solutions from CRM...");

            //EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(_client, false));
            EntityCollection results = Crm.Solution.RetrieveSolutionsFromCrm(_client, false);

            List<CrmSolution> solutions = Class1.CreateCrmSolutionView(results);

            if (selectedSolutionId != Guid.Empty)
            {
                var sel = solutions.FindIndex(s => s.SolutionId == selectedSolutionId);
                if (sel != -1)
                    Solutions.SelectedIndex = sel;
            }

            Solutions.ItemsSource = solutions;

            HideMessage();

            return true;
        }

        private void ShowMessage(string message)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        LockMessage.Content = message;
                        LockOverlay.Visibility = Visibility.Visible;
                    }
                ));
        }

        private void HideMessage()
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }
    }
}
