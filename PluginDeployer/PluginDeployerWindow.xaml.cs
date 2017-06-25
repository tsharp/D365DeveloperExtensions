using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using PluginDeployer.ViewModels;
using SparkleXrm.Tasks;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;
using Task = System.Threading.Tasks.Task;
using Thread = System.Threading.Thread;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace PluginDeployer
{
    public partial class PluginDeployerWindow : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly EnvDTE.Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private bool _isIlMergeInstalled;

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public PluginDeployerWindow()
        {
            InitializeComponent();
            DataContext = this;

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
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            await GetCrmData();

            if (!ConfigFile.ConfigFileExists(_dte.Solution.FullName))
                ConfigFile.CreateConfigFile(ConnPane.OrganizationId, ConnPane.SelectedProject.UniqueName, _dte.Solution.FullName);

            SetButtonState(true);
        }

        private void SetButtonState(bool enabled)
        {
            Publish.IsEnabled = enabled;
            Customizations.IsEnabled = enabled;
            Solutions.IsEnabled = enabled;
            IlMerge.IsEnabled = enabled;
            SpklInstrument.IsEnabled = enabled;
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
            SolutionList.IsEnabled = false;

            SetButtonState(false);
        }

        private async Task GetCrmData()
        {
            ShowMessage("Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

            var solutionTask = GetSolutions();
            var assembliesTask = GetCrmAssemblies();

            await Task.WhenAll(solutionTask, assembliesTask);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);

            if (!solutionTask.Result)
            {
                MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                return;
            }

            if (!assembliesTask.Result)
                MessageBox.Show("Error Retrieving Assemblies. See the Output Window for additional details.");
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            List<CrmSolution> solutions = ModelBuilder.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = solutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private async Task<bool> GetCrmAssemblies()
        {
            EntityCollection results =
                await Task.Run(() => Crm.Assembly.RetrieveAssembliesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Assemblies From CRM", MessageType.Info);

            List<CrmAssembly> assemblies = ModelBuilder.CreateCrmAssemblyView(results);

            CrmAssemblyList.ItemsSource = assemblies;
            CrmAssemblyList.SelectedIndex = 0;

            return true;
        }

        private void ShowMessage(string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.SetStatusBarValue(_dte, "Retrieving assemblies...", (vsStatusAnimation)animation);
                        LockMessage.Content = message;
                        LockOverlay.Visibility = Visibility.Visible;
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
                        LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }

        private void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                ShowMessage("Deploying...", vsStatusAnimation.vsStatusAnimationDeploy);

                if (CrmAssemblyList.SelectedIndex == -1)
                    return;

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy;
                var ctx = new OrganizationServiceContext(service);

                ProjectWorker.BuildProject(ConnPane.SelectedProject);

                using (ctx)
                {
                    PluginRegistraton pluginRegistraton = new PluginRegistraton(service, ctx, new TraceLogger());

                    string path = ProjectWorker.GetAssemblyPath(ConnPane.SelectedProject);

                    CrmAssembly assembly = (CrmAssembly)CrmAssemblyList.SelectedItem;
                    if (assembly.IsWorkflow)
                        pluginRegistraton.RegisterWorkflowActivities(path);
                    else
                        pluginRegistraton.RegisterPlugin(path);
                }
            }
            finally
            {
                HideMessage(vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void AssemblyList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void IlMerge_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ConnPane.SelectedProject == null)
                    return;

                if (!_isIlMergeInstalled)
                {
                    bool installed = IlMergeHandler.Install(_dte, ConnPane.SelectedProject);

                    //CRM Assemblies shouldn't be copied local to prevent merging
                    if (installed)
                        IlMergeHandler.SetReferenceCopyLocal(ConnPane.SelectedProject, false);

                    SetIlMergeTooltip(true);
                    _isIlMergeInstalled = true;
                }
                else
                {
                    bool uninstalled = IlMergeHandler.Uninstall(_dte, ConnPane.SelectedProject);

                    // Reset CRM Assemblies to copy local
                    if (uninstalled)
                        IlMergeHandler.SetReferenceCopyLocal(ConnPane.SelectedProject, true);

                    SetIlMergeTooltip(false);
                    _isIlMergeInstalled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error installing : " + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void SetIlMergeTooltip(bool installed)
        {
            IlMerge.ToolTip = installed ?
                PluginDeployer.Resources.Resource.ILMergeTooltipRemove :
                PluginDeployer.Resources.Resource.ILMergeTooltipEnable;
        }

        private void ProjectAssemblyList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void SpklInstrument_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                ShowMessage("Instrumenting...", vsStatusAnimation.vsStatusAnimationSync);

                if (CrmAssemblyList.SelectedIndex == -1)
                    return;

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy;
                var ctx = new OrganizationServiceContext(service);

                Project project = ConnPane.SelectedProject;
                ProjectWorker.BuildProject(project);

                string path = Path.GetDirectoryName(project.FullName);

                DownloadPluginMetadataTask downloadPluginMetadataTask = new DownloadPluginMetadataTask(ctx, new TraceLogger());
                downloadPluginMetadataTask.Execute(path);
            }
            finally
            {
                HideMessage(vsStatusAnimation.vsStatusAnimationSync);
            }
        }
    }
}