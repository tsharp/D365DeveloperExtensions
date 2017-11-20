using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.DataGrid;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using WebResourceDeployer.Models;
using WebResourceDeployer.ViewModels;
using static CrmDeveloperExtensions2.Core.FileSystem;
using Mapping = WebResourceDeployer.Config.Mapping;
using Task = System.Threading.Tasks.Task;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace WebResourceDeployer
{
    public partial class WebResourceDeployerWindow : INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly FieldInfo _menuDropAlignmentField;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private ObservableCollection<WebResourceItem> _webResourceItems;
        private ObservableCollection<ComboBoxItem> _projectFiles;
        private ObservableCollection<string> _projectFolders;
        private ObservableCollection<WebResourceType> _webResourceTypes;
        private ObservableCollection<FilterTypeName> _filterTypeNames;
        private ObservableCollection<FilterState> _filterStates;

        public ObservableCollection<WebResourceItem> WebResourceItems
        {
            get => _webResourceItems;
            set
            {
                _webResourceItems = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<WebResourceType> WebResourceTypes
        {
            get => _webResourceTypes;
            set
            {
                _webResourceTypes = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<string> ProjectFolders
        {
            get => _projectFolders;
            set
            {
                _projectFolders = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<ComboBoxItem> ProjectFiles
        {
            get => _projectFiles;
            set
            {
                _projectFiles = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterTypeName> FilterTypeNames
        {
            get => _filterTypeNames;
            set
            {
                _filterTypeNames = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterState> FilterStates
        {
            get => _filterStates;
            set
            {
                _filterStates = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public WebResourceDeployerWindow()
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

            #region Fix for Tablet/Touchscreen left-right menu
            _menuDropAlignmentField = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static);
            System.Diagnostics.Debug.Assert(_menuDropAlignmentField != null);
            EnsureStandardPopupAlignment();
            SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
            #endregion
        }

        #region Fix for Tablet/Touchscreen left-right menu
        private void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            EnsureStandardPopupAlignment();
        }
        private void EnsureStandardPopupAlignment()
        {
            if (SystemParameters.MenuDropAlignment && _menuDropAlignmentField != null)
                _menuDropAlignmentField.SetValue(null, false);
        }
        #endregion

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

            //Grid is populated already
            if (WebResourceGrid.ItemsSource != null)
                return;

            if (ConnPane.CrmService != null && ConnPane.CrmService.IsReady)
            {
                SetWindowCaption(gotFocus.Caption);
                LoadData();

                WebResourceTypes =
                    CrmDeveloperExtensions2.Core.Models.WebResourceTypes.GetTypes(ConnPane.CrmService.ConnectedOrgVersion.Major, true);
            }
        }

        private async void LoadData()
        {
            await GetCrmData();
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.SetCaption(currentCaption, ConnPane.CrmService);
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            WebResourceItems = new ObservableCollection<WebResourceItem>();
            ProjectFiles = new ObservableCollection<ComboBoxItem>();
            ProjectFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);
            ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject);
            FilterTypeNames = new ObservableCollection<FilterTypeName>();
            FilterStates = new ObservableCollection<FilterState>();

            SetWindowCaption(_dte.ActiveWindow.Caption);
            LoadData();
            WebResourceTypes =
                CrmDeveloperExtensions2.Core.Models.WebResourceTypes.GetTypes(ConnPane.CrmService.ConnectedOrgVersion.Major, true);

            SolutionList.IsEnabled = true;
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FilterWebResourceItems();
        }

        private void Search_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            FilterWebResourceItems();
        }

        private void AddWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            ObservableCollection<ComboBoxItem> projectsFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;
            NewWebResource newWebResource = new NewWebResource(ConnPane.CrmService, _dte, projectsFiles, solutionId);
            bool? result = newWebResource.ShowModal();

            if (result != true)
                return;

            WebResourceItem defaultItem = ModelBuilder.WebResourceItemFromNew(newWebResource, newWebResource.NewSolutionId);
            defaultItem.PropertyChanged += WebResourceItem_PropertyChanged;
            //Needs to be after setting the property changed event
            defaultItem.BoundFile = newWebResource.NewBoundFile;

            WebResourceItems.Add(defaultItem);

            if (newWebResource.NewSolutionId != ExtensionConstants.DefaultSolutionId)
            {
                WebResourceItem solutionItem = ModelBuilder.WebResourceItemFromNew(newWebResource, ExtensionConstants.DefaultSolutionId);
                solutionItem.PropertyChanged += WebResourceItem_PropertyChanged;
                //Needs to be after setting the property changed event
                solutionItem.BoundFile = newWebResource.NewBoundFile;
                solutionItem.Description = newWebResource.NewDescription;

                WebResourceItems.Add(solutionItem);
            }

            WebResourceItems = new ObservableCollection<WebResourceItem>(WebResourceItems.OrderBy(w => w.Name));

            FilterWebResourceItems();

            WebResourceGrid.ScrollIntoView(defaultItem);
        }

        private void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            List<WebResourceItem> selectedItems = new List<WebResourceItem>();
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;

            //Check for unsaved & missing files
            List<ProjectItem> dirtyItems = new List<ProjectItem>();
            foreach (var selectedItem in WebResourceItems.Where(w => w.Publish && w.SolutionId == solutionId))
            {
                selectedItems.Add(selectedItem);

                ComboBoxItem item = ProjectFiles.FirstOrDefault(c => c.Content.ToString() == selectedItem.BoundFile);
                if (item == null) continue;

                ProjectItem projectItem = (ProjectItem)item.Tag;
                if (!projectItem.IsDirty) continue;

                string filePath = Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + selectedItem.BoundFile.Replace("/", "\\");
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Could not find file: {selectedItem.BoundFile}");
                    return;
                }

                dirtyItems.Add(projectItem);
            }

            if (dirtyItems.Count > 0)
            {
                var result = MessageBox.Show("Save item(s) and publish?", "Unsaved Item(s)", MessageBoxButton.YesNo);
                if (result != MessageBoxResult.Yes) return;

                foreach (var projectItem in dirtyItems)
                    projectItem.Save();
            }

            //Build TypeScript project
            if (selectedItems.Any(p => p.BoundFile.ToUpper().EndsWith("TS")))
            {
                SolutionBuild solutionBuild = _dte.Solution.SolutionBuild;
                solutionBuild.BuildProject(_dte.Solution.SolutionBuild.ActiveConfiguration.Name, ConnPane.SelectedProject.UniqueName, true);
            }

            UpdateWebResources(selectedItems);
        }

        private async void UpdateWebResources(List<WebResourceItem> items)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Updating & Publishing...", vsStatusAnimation.vsStatusAnimationDeploy);

                List<Entity> webResources = new List<Entity>();
                foreach (WebResourceItem webResourceItem in items)
                {
                    Entity webResource = Crm.WebResource.CreateUpdateWebResourceEntity(webResourceItem.WebResourceId,
                        webResourceItem.BoundFile, webResourceItem.Description, ConnPane.SelectedProject.FullName);
                    webResources.Add(webResource);
                }

                bool success;
                //Check if < CRM 2011 UR12 (ExecuteMutliple)
                Version version = ConnPane.CrmService.ConnectedOrgVersion;
                if (version.Major == 5 && version.Revision < 3200)
                    success = await Task.Run(() => Crm.WebResource.UpdateAndPublishSingle(ConnPane.CrmService, webResources));
                else
                    success = await Task.Run(() => Crm.WebResource.UpdateAndPublishMultiple(ConnPane.CrmService, webResources));

                if (success) return;

                MessageBox.Show("Error Updating And Publishing Web Resources. See the Output Window for additional details.");
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void WebResourceGrid_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Make rows unselectable
            WebResourceGrid.UnselectAllCells();
        }

        private void ConnPane_SelectedProjectChanged(object sender, SelectionChangedEventArgs e)
        {
            if (WebResourceItems != null)
                ResetForm();
        }

        private void ConnPane_SolutionProjectAdded(object sender, SolutionProjectAddedEventArgs e)
        {

        }

        private void ConnPane_OnSolutionProjectRemoved(object sender, SolutionProjectRemovedEventArgs e)
        {
            Project project = e.Project;
            if (ConnPane.SelectedProject == project)
                ResetForm();
        }

        private void UpdateWebResourceItemsBoundFile(string oldValue, string newValue)
        {
            foreach (WebResourceItem webResourceItem in WebResourceItems
                .Where(w => w.BoundFile != null && w.BoundFile.Equals(oldValue, StringComparison.InvariantCultureIgnoreCase)))
            {
                webResourceItem.BoundFile = newValue;
            }
        }

        private void UpdateProjectFilesAfterChange(string oldName, string newName)
        {
            ComboBoxItem projectFile = ProjectFiles.FirstOrDefault(p => p.Content.ToString().Equals(oldName, StringComparison.InvariantCultureIgnoreCase));
            if (projectFile == null)
                return;

            if (string.IsNullOrEmpty(newName))
                ProjectFiles.Remove(projectFile);
            else
                projectFile.Content = newName;
        }

        private void UpdateWebResourceItemsBoundFilePath(string oldName, string newName)
        {
            foreach (WebResourceItem webResourceItem in WebResourceItems.Where(
                w => w.BoundFile != null && w.BoundFile.StartsWith(oldName)))
            {
                webResourceItem.BoundFile = string.IsNullOrEmpty(newName)
                    ? null
                    : webResourceItem.BoundFile.Replace(oldName, newName);
            }
        }

        private void UpdateProjectFilesPathsAfterChange(string oldItemName, string newItemPath)
        {
            List<ComboBoxItem> projectFilesToRename = ProjectFiles.Where(p => p.Content.ToString().StartsWith(oldItemName)).ToList();
            foreach (ComboBoxItem comboBoxItem in projectFilesToRename)
            {
                if (string.IsNullOrEmpty(newItemPath))
                    ProjectFiles.Remove(comboBoxItem);
                else
                    comboBoxItem.Content = comboBoxItem.Content.ToString().Replace(oldItemName, newItemPath);
            }
        }

        private void ConnPane_OnProjectItemRenamed(object sender, ProjectItemRenamedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;
            if (projectItem.Name == null) return;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            string oldName = e.OldName;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string newItemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);
                string oldItemName = newItemName.Replace(Path.GetFileName(projectItem.Name), oldName).Replace("//", "/");

                UpdateWebResourceItemsBoundFile(oldItemName, newItemName);

                UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = LocalPathToCrmPath(projectPath, projectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

                UpdateWebResourceItemsBoundFilePath(oldItemPath, newItemPath);

                UpdateProjectFilesPathsAfterChange(oldItemPath, newItemPath);
            }
        }

        private void ConnPane_OnProjectItemAdded(object sender, ProjectItemAddedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;

            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                if (projectPath == null) return;

                string newItemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                if (ProjectFiles == null)
                    ProjectFiles = new ObservableCollection<ComboBoxItem>();

                ProjectFiles.Add(new ComboBoxItem
                {
                    Content = newItemName
                });

                ProjectFiles = new ObservableCollection<ComboBoxItem>(ProjectFiles.OrderBy(p => p.Content.ToString()));
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
                ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject);
        }

        private void ConnPane_OnProjectItemRemoved(object sender, ProjectItemRemovedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return;

            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string itemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                UpdateWebResourceItemsBoundFile(itemName, null);

                UpdateProjectFilesAfterChange(itemName, null);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var itemName = LocalPathToCrmPath(projectPath,
                    projectItem.FileNames[1].TrimEnd(Path.DirectorySeparatorChar));

                int index = itemName.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                ProjectFolders.Remove(itemName);

                UpdateWebResourceItemsBoundFilePath(itemName, null);

                UpdateProjectFilesPathsAfterChange(itemName, null);
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
                string newItemName = LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1]);

                UpdateWebResourceItemsBoundFile(oldItemName, newItemName);

                UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = newItemPath.LastIndexOf(postMoveProjectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                UpdateWebResourceItemsBoundFilePath(oldItemName, newItemPath);

                UpdateProjectFilesPathsAfterChange(oldItemName, newItemPath);

                ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject);
            }
        }

        private void SetButtonState(bool enabled)
        {
            Publish.IsEnabled = enabled;
        }

        private void ResetForm()
        {
            ProjectFiles = new ObservableCollection<ComboBoxItem>();
            ProjectFolders = new ObservableCollection<string>();
            WebResourceItems = new ObservableCollection<WebResourceItem>();
            FilterTypeNames = new ObservableCollection<FilterTypeName>();
            FilterStates = new ObservableCollection<FilterState>();

            SolutionList.IsEnabled = false;
            SolutionList.ItemsSource = null;
            WebResourceGrid.IsEnabled = false;
            Search.Text = String.Empty;
            Search.IsEnabled = false;
            SetButtonState(false);

            ClearConnection();
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();

            ClearConnection();
        }

        private void ConnPane_OnSolutionOpened(object sender, EventArgs e)
        {
            ClearConnection();
        }

        private void ClearConnection()
        {
            ConnPane.IsConnected = false;
            ConnPane.CrmService?.Dispose();
            ConnPane.CrmService = null;
        }

        private void PublishSelectAll_OnClick(object sender, RoutedEventArgs e)
        {
            CheckBox publishAll = (CheckBox)sender;
            bool? isChecked = publishAll.IsChecked;

            if (isChecked != null && isChecked.Value)
                UpdateAllPublishChecks(true);
            else
                UpdateAllPublishChecks(false);
        }

        private void UpdateAllPublishChecks(bool publish)
        {
            foreach (WebResourceItem webResourceItem in WebResourceItems)
                if (webResourceItem.AllowPublish)
                    webResourceItem.Publish = publish;
        }

        private void BoundFile_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Grid grid = (Grid)sender;
            TextBlock textBlock = (TextBlock)grid.Children[0];

            Guid webResourceId = new Guid(textBlock.Tag.ToString());
            FileId.Content = webResourceId;

            WebResourceItem webResourceItem = WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);
            ProjectFileList.SelectedIndex = -1;
            if (webResourceItem != null)
            {
                foreach (ComboBoxItem comboBoxItem in ProjectFileList.Items)
                {
                    if (comboBoxItem.Content.ToString() != webResourceItem.BoundFile) continue;

                    ProjectFileList.SelectedItem = comboBoxItem;
                    break;
                }
            }

            //Fixes the placement of the popup so it fits in the cell w/ the padding applied
            ProjectFileList.Width = WebResourceGrid.Columns[6].ActualWidth - 1;
            FilePopup.PlacementTarget = grid;
            FilePopup.Placement = PlacementMode.Relative;
            FilePopup.VerticalOffset = -3;
            FilePopup.HorizontalOffset = -2;

            FilePopup.IsOpen = true;
            ProjectFileList.IsDropDownOpen = true;
        }

        private async void GetWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
            WebResourceItem webResourceItem =
                WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);

            string folder = String.Empty;
            if (!string.IsNullOrEmpty(webResourceItem?.BoundFile))
                folder = Crm.WebResource.GetExistingFolderFromBoundFile(webResourceItem, folder);

            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService);
        }

        private async void DownloadWebResourceToFolder(string folder, Guid webResourceId)
        {
            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService);
        }

        private void DownloadAll_OnClick(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("This will download all web resources in the current view to a matching folder structure" +
                Environment.NewLine + Environment.NewLine + "OK to proceed?", "Web Resource Download",
                MessageBoxButton.YesNo);
            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                Overlay.ShowMessage(_dte, "Downloading file(s)...", vsStatusAnimation.vsStatusAnimationSync);

                ICollectionView icv = CollectionViewSource.GetDefaultView(WebResourceGrid.ItemsSource);
                List<WebResourceItem> downloadItems = icv.Cast<WebResourceItem>().ToList();

                Dictionary<Guid, string> updatesToMake = new Dictionary<Guid, string>();

                Parallel.ForEach(downloadItems, currentItem =>
                {
                    Entity webResource = Crm.WebResource.RetrieveWebResourceFromCrm(ConnPane.CrmService, currentItem.WebResourceId);

                    OutputLogger.WriteToOutputWindow("Downloaded Web Resource: " + webResource.Id, MessageType.Info);

                    string name = webResource.GetAttributeValue<string>("name");
                    name = Crm.WebResource.AddMissingExtension(name, webResource.GetAttributeValue<OptionSetValue>("webresourcetype").Value);

                    //TODO: option to change root folder
                    string path = Crm.WebResource.ConvertWebResourceNameFullToPath(name, "/", ConnPane.SelectedProject);

                    string content = Crm.WebResource.GetWebResourceContent(webResource);
                    byte[] decodedContent = Crm.WebResource.DecodeWebResource(content);

                    WriteFileToDisk(path, decodedContent);

                    ProjectItem projectItem = ConnPane.SelectedProject.ProjectItems.AddFromFile(path);

                    var fullname = projectItem.FileNames[1];
                    var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                    if (projectPath == null) return;

                    var boundName = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                    updatesToMake.Add(currentItem.WebResourceId, boundName);
                });

                foreach (var update in updatesToMake)
                {
                    foreach (WebResourceItem item in WebResourceItems.Where(w => w.WebResourceId == update.Key))
                    {
                        item.BoundFile = update.Value;
                    }
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task DownloadWebResourceAsync(Guid webResourceId, string folder, CrmServiceClient client)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Downloading file...", vsStatusAnimation.vsStatusAnimationSync);

                Entity webResource = await Task.Run(() => Crm.WebResource.RetrieveWebResourceFromCrm(client, webResourceId));

                OutputLogger.WriteToOutputWindow("Downloaded Web Resource: " + webResource.Id, MessageType.Info);

                string name = webResource.GetAttributeValue<string>("name");
                name = Crm.WebResource.AddMissingExtension(name, webResource.GetAttributeValue<OptionSetValue>("webresourcetype").Value);

                string path = Crm.WebResource.ConvertWebResourceNameToPath(name, folder, ConnPane.SelectedProject.FullName);

                if (File.Exists(path))
                {
                    MessageBoxResult result = MessageBox.Show("OK to overwrite?", "Web Resource Download",
                        MessageBoxButton.YesNo);
                    if (result != MessageBoxResult.Yes)
                    {
                        Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                        return;
                    }
                }

                string content = Crm.WebResource.GetWebResourceContent(webResource);
                byte[] decodedContent = Crm.WebResource.DecodeWebResource(content);

                WriteFileToDisk(path, decodedContent);

                ProjectItem projectItem = ConnPane.SelectedProject.ProjectItems.AddFromFile(path);

                var fullname = projectItem.FileNames[1];
                var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                if (projectPath == null) return;

                var boundName = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                foreach (WebResourceItem item in WebResourceItems.Where(w => w.WebResourceId == webResourceId))
                {
                    item.BoundFile = boundName;

                    CheckBox publishAll =
                        DataGridHelpers.FindVisualChildren<CheckBox>(WebResourceGrid)
                            .FirstOrDefault(t => t.Name == "PublishSelectAll");
                    if (publishAll == null) return;

                    if (publishAll.IsChecked == true)
                        item.Publish = true;
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void OpenWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());

            string contentUrl = $"main.aspx?etc=9333&id=%7b{webResourceId}%7d&pagetype=webresourceedit";

            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService, contentUrl);
        }

        private async void CompareWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Downloading file for compare...", vsStatusAnimation.vsStatusAnimationSync);

                //Get the file from CRM and save in temp files
                Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
                Entity webResource = await Task.Run(() => Crm.WebResource.RetrieveWebResourceContentFromCrm(ConnPane.CrmService, webResourceId));

                OutputLogger.WriteToOutputWindow($"Retrieved Web Resource {webResourceId} For Compare", MessageType.Info);

                string tempFile = WriteTempFile(webResource.GetAttributeValue<string>("name"),
                        Crm.WebResource.DecodeWebResource(webResource.GetAttributeValue<string>("content")));

                Project project = ConnPane.SelectedProject;
                var projectPath = Path.GetDirectoryName(project.FullName);
                if (projectPath == null) return;

                var webResourceItem = WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);
                if (webResourceItem == null)
                    return;

                string boundFilePath = webResourceItem.BoundFile;

                _dte.ExecuteCommand("Tools.DiffFiles",
                    $"\"{tempFile}\" \"{projectPath + boundFilePath.Replace("/", "\\")}\" \"{webResource.GetAttributeValue<string>("name") + " - CRM"}\" \"{boundFilePath + " - Local"}\"");
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async void DeleteWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            MessageBoxResult deleteResult = MessageBox.Show("Are you sure?" + Environment.NewLine + Environment.NewLine +
                                                            "This will attempt to delete the web resource from CRM.", "Delete Web Resource", MessageBoxButton.YesNo);
            if (deleteResult != MessageBoxResult.Yes)
                return;

            try
            {
                Overlay.ShowMessage(_dte, "Deleting web resource...", vsStatusAnimation.vsStatusAnimationSync);

                Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
                await Task.Run(() => Crm.WebResource.DeleteWebResourcetFromCrm(ConnPane.CrmService, webResourceId));

                foreach (WebResourceItem webResourceItem in WebResourceItems.Where(w => w.WebResourceId == webResourceId))
                    WebResourceItems.Remove(webResourceItem);

                WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

                FilterWebResourceItems();

                OutputLogger.WriteToOutputWindow($"Deleted Web Resource {webResourceId}", MessageType.Info);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void ProjectFileList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectFileList.SelectedIndex == -1) return;

            WebResourceItem webResourceItem =
                WebResourceItems.FirstOrDefault(w => w.WebResourceId == new Guid(FileId.Content.ToString()));

            ComboBoxItem item = (ComboBoxItem)ProjectFileList.SelectedItem;

            if (webResourceItem != null && webResourceItem.BoundFile != item.Content.ToString())
                webResourceItem.BoundFile = item.Content.ToString();

            FilePopup.IsOpen = false;
        }

        private async Task GetCrmData()
        {
            try
            {
                Overlay.ShowMessage(_dte, "Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();
                var webResourceTask = GetWebResources();

                await Task.WhenAll(solutionTask, webResourceTask);

                if (!solutionTask.Result)
                {
                    MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                    return;
                }

                if (!webResourceTask.Result)
                    MessageBox.Show("Error Retrieving Web Resources. See the Output Window for additional details.");
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService, true));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            List<CrmSolution> solutions = ModelBuilder.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = solutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private async Task<bool> GetWebResources()
        {
            EntityCollection results = await Task.Run(() => Crm.WebResource.RetrieveWebResourcesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Web Resources From CRM", MessageType.Info);

            WebResourceItems = ModelBuilder.CreateWebResourceItemView(results, ConnPane.SelectedProject.Name);

            WebResourceItems = new ObservableCollection<WebResourceItem>(WebResourceItems.OrderBy(w => w.Name));

            CreateFilterTypeNameList();
            CreateFilterStateList();

            WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

            foreach (WebResourceItem webResourceItem in WebResourceItems)
            {
                webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
            }

            WebResourceGrid.IsEnabled = true;
            Search.IsEnabled = true;
            FilterWebResourceItems();

            return true;
        }

        private void WebResourceItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            WebResourceItem item = (WebResourceItem)sender;
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                              ExtensionConstants.DefaultSolutionId;

            if (e.PropertyName == "BoundFile")
            {
                if (WebResourceGrid.ItemsSource != null)
                    foreach (WebResourceItem webResourceItem in WebResourceItems.Where(w => w.Name == item.Name && w.BoundFile != item.BoundFile))
                    {
                        //Remove & re-add the event handler to avoid triggering on same file on other solutions
                        webResourceItem.PropertyChanged -= WebResourceItem_PropertyChanged;
                        webResourceItem.BoundFile = item.BoundFile;
                        webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
                        if (string.IsNullOrEmpty(item.BoundFile) && item.Publish)
                            webResourceItem.Publish = false;
                    }

                Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, item);
            }

            if (e.PropertyName == "Publish")
            {

                foreach (WebResourceItem webResourceItem in WebResourceItems.Where(w => w.WebResourceId == item.WebResourceId && w.SolutionId == solutionId))
                    webResourceItem.Publish = item.Publish;

                Publish.IsEnabled = WebResourceItems.Count(w => w.Publish) > 0;

                SetPublishAll();
            }
        }

        private void SetPublishAll()
        {
            CheckBox publishAll = DataGridHelpers.FindVisualChildren<CheckBox>(WebResourceGrid).FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll == null) return;

            publishAll.IsChecked = WebResourceItems.Count(w => w.Publish) == WebResourceItems.Count(w => w.AllowPublish);
        }

        private void CreateFilterTypeNameList()
        {
            FilterTypeNames = FilterTypeName.CreateFilterList(WebResourceItems);

            foreach (FilterTypeName filterTypeName in FilterTypeNames)
            {
                filterTypeName.PropertyChanged += FilterTypeName_PropertyChanged;
            }

            FilterByTypeName.IsEnabled = FilterTypeNames.Count > 2;
        }

        private void CreateFilterStateList()
        {
            FilterStates = FilterState.CreateFilterList(WebResourceItems);

            foreach (FilterState filterTypeState in FilterStates)
            {
                filterTypeState.PropertyChanged += FilterState_PropertyChanged;
            }

            FilterByState.IsEnabled = FilterStates.Count > 2;
        }

        private void FilterTypeName_PropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            GridFilters.SetSelectAll(sender, FilterTypeNames);

            FilterWebResourceItems();
        }

        private void FilterState_PropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            GridFilters.SetSelectAll(sender, FilterStates);

            FilterWebResourceItems();
        }

        public void FilterWebResourceItems()
        {
            if (WebResourceItems.Count == 0)
                return;

            ICollectionView icv = CollectionViewSource.GetDefaultView(WebResourceGrid.ItemsSource);
            if (icv == null) return;

            icv.Filter = GetFilteredView;

            //Only keep publish flags for items still visable
            foreach (WebResourceItem item in WebResourceItems.Where(w => w.Publish).Except(icv.OfType<WebResourceItem>()))
            {
                item.Publish = false;
            }
        }

        public bool GetFilteredView(object sourceObject)
        {
            WebResourceItem webResourceItem = (WebResourceItem)sourceObject;

            if (TypeNameCondition(webResourceItem) && StateCondition(webResourceItem)
                && SolutionCondition(webResourceItem) && SearchNameCondition(webResourceItem))
            {
                return true;
            }

            return false;
        }

        public bool SearchNameCondition(WebResourceItem webResourceItem)
        {
            string search = Search.Text.Trim();
            if (search.Length < 2)
                return true;


            if (!string.IsNullOrEmpty(webResourceItem.Name))
            {
                if (webResourceItem.Name.Contains(search))
                    return true;
            }

            if (!string.IsNullOrEmpty(webResourceItem.DisplayName))
            {
                if (webResourceItem.DisplayName.Contains(search))
                    return true;
            }

            if (!string.IsNullOrEmpty(webResourceItem.Description))
            {
                if (webResourceItem.Description.Contains(search))
                    return true;
            }

            return false;
        }

        public bool SolutionCondition(WebResourceItem webResourceItem)
        {
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
              ExtensionConstants.DefaultSolutionId;

            return webResourceItem.SolutionId == solutionId;
        }

        public bool TypeNameCondition(WebResourceItem webResourceItem)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterTypeNames), webResourceItem.TypeName);
        }

        public bool StateCondition(WebResourceItem webResourceItem)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterStates), webResourceItem.State);
        }

        public bool IsStringFilterValid(ObservableCollection<IFilterProperty> list, string name)
        {
            return list.Count(e => e.IsSelected) != 0 &&
                   list.Where(e => e.IsSelected).Select(e => e.Name).Contains(name);
        }

        private void Refresh_OnClick(object sender, RoutedEventArgs e)
        {
            RefreshFromCrm();
        }

        private async void RefreshFromCrm()
        {
            try
            {
                Overlay.ShowMessage(_dte, "Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

                var toPublish = WebResourceItems.Where(w => w.Publish);

                EntityCollection results = await Task.Run(() => Crm.WebResource.RetrieveWebResourcesFromCrm(ConnPane.CrmService));
                if (results == null)
                {
                    Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Retrieving Web Resources. See the Output Window for additional details.");
                    return;
                }

                OutputLogger.WriteToOutputWindow("Retrieved Web Resources From CRM", MessageType.Info);

                WebResourceItems = ModelBuilder.CreateWebResourceItemView(results, ConnPane.SelectedProject.Name);

                foreach (WebResourceItem webResourceItem in WebResourceItems)
                {
                    webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
                }

                WebResourceItems = new ObservableCollection<WebResourceItem>(WebResourceItems.OrderBy(w => w.Name));

                WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

                FilterWebResourceItems();

                var toUpdate = WebResourceItems.Where(w => toPublish.Any(t => w.WebResourceId == t.WebResourceId));
                foreach (WebResourceItem item in toUpdate)
                {
                    item.Publish = true;
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void ProjectFolderList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox projectFolderList = (ComboBox)sender;
            string folder = projectFolderList.SelectedValue?.ToString();
            if (string.IsNullOrEmpty(folder)) return;
            Guid webResourceId = new Guid(FolderId.Content.ToString());
            DownloadWebResourceToFolder(folder, webResourceId);
        }

        private void GetWebResource_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            Button button = (Button)sender;
            FolderId.Content = button.CommandParameter;

            FolderPopup.PlacementTarget = button;
            FolderPopup.Placement = PlacementMode.Relative;

            FolderPopup.IsOpen = true;
            ProjectFolderList.IsDropDownOpen = true;
        }

        private void FilterByTypeName_Click(object sender, RoutedEventArgs e)
        {
            FilterTypeNamePopup.OpenFilterList(sender);
        }

        private void FilterByState_Click(object sender, RoutedEventArgs e)
        {
            FilterStatePopup.OpenFilterList(sender);
        }

        private static void ShowHideDescriptionButtons(object sender, Visibility editVisibility, Visibility saveCancelVisibility)
        {
            Button editDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "EditDescription");
            editDescription.Visibility = editVisibility;

            Button cancelEditDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "CancelEditDescription");
            cancelEditDescription.Visibility = saveCancelVisibility;

            Button saveDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "SaveDescription");
            saveDescription.Visibility = saveCancelVisibility;

            Button undoEditDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "UndoEditDescription");
            undoEditDescription.Visibility = saveCancelVisibility;
        }

        private void EditDescription_Click(object sender, RoutedEventArgs e)
        {
            ShowHideDescriptionButtons(sender, Visibility.Hidden, Visibility.Visible);
            DetailsRow.ShowHideDetailsRow(sender);
        }

        private void CancelEditDescription_Click(object sender, RoutedEventArgs e)
        {
            ShowHideDescriptionButtons(sender, Visibility.Visible, Visibility.Hidden);
            DetailsRow.ShowHideDetailsRow(sender);

            var webResourceItem = WebResourceItemFromCmdParam(sender);
            webResourceItem.PreviousDescription = webResourceItem.Description;

            foreach (WebResourceItem item in WebResourceItems.Where(w => w.WebResourceId == webResourceItem.WebResourceId))
            {
                item.Description = item.PreviousDescription;
            }
        }

        private WebResourceItem WebResourceItemFromCmdParam(object sender)
        {
            Button button = (Button)sender;
            Guid webResourceId = (Guid)button.CommandParameter;
            WebResourceItem webResourceItem = WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);
            return webResourceItem;
        }

        private void SaveDescription_Click(object sender, RoutedEventArgs e)
        {
            ShowHideDescriptionButtons(sender, Visibility.Visible, Visibility.Hidden);
            DetailsRow.ShowHideDetailsRow(sender);

            WebResourceItem webResourceItem = WebResourceItemFromCmdParam(sender);

            TextBox description = DetailsRow.GetDataGridRowControl<TextBox>(sender, "Description");
            string newDescription = description.Text.Trim();

            if (webResourceItem.PreviousDescription == newDescription)
                return;

            webResourceItem.Description = newDescription;
            webResourceItem.PreviousDescription = newDescription;

            SetWebResourceItemDescriptions(webResourceItem.WebResourceId, newDescription);

            List<Entity> webResources = new List<Entity>();
            Entity webResource = new Entity("webresource")
            {
                Id = webResourceItem.WebResourceId,
                ["description"] = newDescription
            };

            webResources.Add(webResource);

            ExecuteSaveDescription(webResources, webResourceItem);
        }

        private async void ExecuteSaveDescription(List<Entity> webResources, WebResourceItem webResourceItem)
        {
            await Task.Run(() => SaveDescription(webResources, webResourceItem));

            Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, webResourceItem);
        }

        private void SaveDescription(List<Entity> webResources, WebResourceItem webResourceItem)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Updating", vsStatusAnimation.vsStatusAnimationDeploy);

                Crm.WebResource.UpdateAndPublishSingle(ConnPane.CrmService, webResources);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void UndoEditDescription_Click(object sender, RoutedEventArgs e)
        {
            ShowHideDescriptionButtons(sender, Visibility.Visible, Visibility.Hidden);
            DetailsRow.ShowHideDetailsRow(sender);

            WebResourceItem webResourceItem = WebResourceItemFromCmdParam(sender);

            ExecuteUndoDescription(webResourceItem);
        }

        private async void ExecuteUndoDescription(WebResourceItem webResourceItem)
        {
            string serverDescription = await Task.Run(() => UndoDescription(webResourceItem));

            SetWebResourceItemDescriptions(webResourceItem.WebResourceId, serverDescription);

            Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, webResourceItem);
        }

        private string UndoDescription(WebResourceItem webResourceItem)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Updating", vsStatusAnimation.vsStatusAnimationDeploy);

                string serverDescription =
                    Crm.WebResource.RetrieveWebResourceDescriptionFromCrm(ConnPane.CrmService, webResourceItem.WebResourceId);

                return serverDescription;
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void SetWebResourceItemDescriptions(Guid webResourceId, string description)
        {
            foreach (WebResourceItem item in WebResourceItems.Where(w => w.WebResourceId == webResourceId))
            {
                item.Description = description;
                item.PreviousDescription = description;
            }
        }

        private void ShowId_Click(object sender, RoutedEventArgs e)
        {
            var width = WebResourceGrid.Columns[0].Width;

            int scaleX = -1;
            if (width.UnitType == DataGridLengthUnitType.SizeToCells)
            {
                WebResourceGrid.Columns[0].Width = new DataGridLength(25, DataGridLengthUnitType.Pixel);
                scaleX = 1;
            }
            else
            {
                WebResourceGrid.Columns[0].Width = DataGridLength.SizeToCells;
                scaleX = -1;
            }

            Button showId = (Button)sender;
            showId.RenderTransformOrigin = new Point(0.5, 0.5);
            ScaleTransform flipTrans = new ScaleTransform
            {
                ScaleX = scaleX
            };
            showId.RenderTransform = flipTrans;
        }

        private void ConnPane_ProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ConnPane.CrmService != null && ConnPane.CrmService.IsReady)
            {
                foreach (WebResourceItem webResourceItem in WebResourceItems)
                {
                    webResourceItem.PropertyChanged -= WebResourceItem_PropertyChanged;
                    webResourceItem.BoundFile = null;
                    webResourceItem.Description = null;
                    webResourceItem.PreviousDescription = null;
                    webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
                }
                WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);
            }
        }
    }
}