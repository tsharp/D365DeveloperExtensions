using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Controls;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using NLog;
using SolutionPackager.ViewModels;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;
using Task = System.Threading.Tasks.Task;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace SolutionPackager
{
    public partial class SolutionPackagerWindow : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();

        private ObservableCollection<CrmSolution> _solutionData;
        public ObservableCollection<CrmSolution> SolutionData
        {
            get => _solutionData;
            set
            {
                _solutionData = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public SolutionPackagerWindow()
        {
            InitializeComponent();
            DataContext = this;

            //TODO: would be better if this used a converter in xaml
            Customizations.Content = Customizations.Content.ToString().ToUpper();
            Solutions.Content = Solutions.Content.ToString().ToUpper();
            PackageSolution.Content = PackageSolution.Content.ToString().ToUpper();
            UnpackageSolution.Content = UnpackageSolution.Content.ToString().ToUpper();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
        }

        private void WindowEventsOnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //WindowEventsOnWindowActivated in this project can be called when activating another window
            //so we don't want to contine further unless our window is active
            if (!HostWindow.IsCrmDevExWindow(gotFocus))
                return;

            SetWindowCaption(gotFocus.Caption);
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.SetCaption(currentCaption, ConnPane.CrmService);
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            SetButtonState(true);

            await GetCrmData();

            if (!ConfigFile.ConfigFileExists(_dte.Solution.FullName))
                ConfigFile.CreateConfigFile(ConnPane.OrganizationId, ConnPane.SelectedProject.UniqueName, _dte.Solution.FullName);
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();
        }

        private void ConnPane_OnSolutionProjectRemoved(object sender, SolutionProjectRemovedEventArgs e)
        {
            Project project = e.Project;
            if (ConnPane.SelectedProject == project)
                ResetForm();
        }

        private void ConnPane_OnSolutionProjectRenamed(object sender, SolutionProjectRenamedEventArgs e)
        {
            Project project = e.Project;
            string solutionPath = Path.GetDirectoryName(_dte.Solution.FullName);
            if (string.IsNullOrEmpty(solutionPath))
                return;

            string oldName = e.OldName.Replace(solutionPath, string.Empty).Substring(1);

            //Mapping.UpdateProjectName(_dte.Solution.FullName, oldName, project.UniqueName);
        }

        private void SetButtonState(bool enabled)
        {
            Customizations.IsEnabled = enabled;
            Solutions.IsEnabled = enabled;
            PackageSolution.IsEnabled = enabled;
            UnpackageSolution.IsEnabled = enabled;
            SolutionList.IsEnabled = enabled;
            DownloadManaged.IsEnabled = enabled;
        }

        private void ConnPane_OnProjectItemRenamed(object sender, ProjectItemRenamedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.Name == null) return;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            string oldName = e.OldName;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string newItemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);
                string oldItemName = newItemName.Replace(Path.GetFileName(projectItem.Name), oldName).Replace("//", "/");

                //UpdateWebResourceItemsBoundFile(oldItemName, newItemName);

                //UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

                //UpdateWebResourceItemsBoundFilePath(oldItemPath, newItemPath);

                //UpdateProjectFilesPathsAfterChange(oldItemPath, newItemPath);
            }
        }

        private void ConnPane_OnProjectItemRemoved(object sender, ProjectItemRemovedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string itemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                //UpdateWebResourceItemsBoundFile(itemName, null);

                //UpdateProjectFilesAfterChange(itemName, null);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var itemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = itemName.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                //UpdateWebResourceItemsBoundFilePath(itemName, null);

                //UpdateProjectFilesPathsAfterChange(itemName, null);
            }
        }

        private void ConnPane_OnProjectItemMoved(object sender, ProjectItemMovedEventArgs e)
        {
            ProjectItem postMoveProjectItem = e.PostMoveProjectItem;
            string oldItemName = e.PreMoveName;
            var projectPath = Path.GetDirectoryName(postMoveProjectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            Guid itemType = new Guid(postMoveProjectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string newItemName = FileSystem.LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1]);

                //UpdateWebResourceItemsBoundFile(oldItemName, newItemName);

                //UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = FileSystem.LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = newItemPath.LastIndexOf(postMoveProjectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                //UpdateWebResourceItemsBoundFilePath(oldItemName, newItemPath);

                //UpdateProjectFilesPathsAfterChange(oldItemName, newItemPath);
            }
        }
        private void ConnPane_OnProjectItemAdded(object sender, ProjectItemAddedEventArgs e)
        {
        }

        private void ResetForm()
        {
            SolutionList.ItemsSource = null;
            SetButtonState(false);
        }

        private async Task GetCrmData()
        {
            try
            {
                ShowMessage("Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();

                await Task.WhenAll(solutionTask);

                if (!solutionTask.Result)
                {
                    HideMessage(vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                }
            }
            finally
            {
                HideMessage(vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            SolutionData = ModelBuilder.CreateCrmSolutionView(results);
            SolutionList.DisplayMemberPath = "NameVersion";

            return true;
        }

        private void ShowMessage(string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.SetStatusBarValue(_dte, "Retrieving web resources...", (vsStatusAnimation)animation);
                        Overlay.Show(message);
                    }
                ));
        }

        private void HideMessage(vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.ClearStatusBarValue(_dte, (vsStatusAnimation)animation);
                        Overlay.Hide();
                    }
                ));
        }

        private void Publish_OnClick(object sender, RoutedEventArgs e)
        {
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService,
                $"tools/solution/edit.aspx?id=%7b{ExtensionConstants.DefaultSolutionId}%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService,
                "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void DownloadManaged_OnChecked(object sender, RoutedEventArgs e)
        {
        }

        private void PackageSolution_OnClick(object sender, RoutedEventArgs e)
        {
        }

        private void UnpackageSolution_OnClick(object sender, RoutedEventArgs e)
        {
        }
    }
}