using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.DataGrid;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.ExtensionMethods;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Vs;
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
using System.Windows.Data;
using System.Windows.Input;
using WebResourceDeployer.Crm;
using WebResourceDeployer.Models;
using WebResourceDeployer.Resources;
using WebResourceDeployer.ViewModels;
using static D365DeveloperExtensions.Core.FileSystem;
using CoreWebResourceTypes = D365DeveloperExtensions.Core.Models.WebResourceTypes;
using Mapping = WebResourceDeployer.Config.Mapping;
using Solution = EnvDTE.Solution;
using Task = System.Threading.Tasks.Task;
using WebBrowser = D365DeveloperExtensions.Core.WebBrowser;

namespace WebResourceDeployer
{
    public partial class WebResourceDeployerWindow : INotifyPropertyChanged
    {
        #region Private

        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly FieldInfo _menuDropAlignmentField;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private ObservableCollection<WebResourceItem> _webResourceItems;
        private ObservableCollection<ComboBoxItem> _projectFiles;
        private ObservableCollection<string> _projectFolders;
        private ObservableCollection<WebResourceType> _webResourceTypes;
        private ObservableCollection<FilterTypeName> _filterTypeNames;
        private ObservableCollection<FilterState> _filterStates;
        private List<MovedWebResourceItem> _movedWebResourceItems;

        #endregion

        #region Public

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

        #endregion

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

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
            //so we don't want to continue further unless our window is active
            if (!HostWindow.IsD365DevExWindow(gotFocus))
                return;

            //Grid is populated already
            if (WebResourceGrid.ItemsSource != null)
                return;

            if (ConnPane.CrmService?.IsReady == true)
                InitializeForm();
        }

        private void ResetCollections()
        {
            ProjectFiles = new ObservableCollection<ComboBoxItem>();
            ProjectFolders = new ObservableCollection<string>();
            WebResourceItems = new ObservableCollection<WebResourceItem>();
            FilterTypeNames = new ObservableCollection<FilterTypeName>();
            FilterStates = new ObservableCollection<FilterState>();
        }

        private async void LoadData()
        {
            ConnPane.CollapsePane();

            await GetCrmData();
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.GetCaption(currentCaption, ConnPane.CrmService);
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            InitializeForm();
        }

        private void InitializeForm()
        {
            ResetCollections();

            var hasTsConfig = TsHelper.HasTsConfig(ConnPane.SelectedProject);
            ProjectFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject, hasTsConfig);
            ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject, ProjectType.WebResource);

            SetWindowCaption(_dte.ActiveWindow.Caption);
            LoadData();
            WebResourceTypes = CoreWebResourceTypes.GetTypes(ConnPane.CrmService.ConnectedOrgVersion.Major, true);

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
            var hasTsConfig = TsHelper.HasTsConfig(ConnPane.SelectedProject);
            var projectsFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject, hasTsConfig);
            var solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;
            var newWebResource = new NewWebResource(ConnPane.CrmService, _dte, projectsFiles, solutionId, ConnPane.SelectedProject);
            var result = newWebResource.ShowModal();

            if (result != true)
                return;

            var defaultItem = ModelBuilder.WebResourceItemFromNew(newWebResource, newWebResource.NewSolutionId);
            defaultItem.PropertyChanged += WebResourceItem_PropertyChanged;
            defaultItem.BoundFile = newWebResource.NewBoundFile;

            WebResourceItems.Add(defaultItem);

            if (newWebResource.NewSolutionId != ExtensionConstants.DefaultSolutionId)
            {
                var solutionItem = ModelBuilder.WebResourceItemFromNew(newWebResource, ExtensionConstants.DefaultSolutionId);
                solutionItem.PropertyChanged += WebResourceItem_PropertyChanged;
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
            var selectedItems = new List<WebResourceItem>();
            var solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;

            //Check for missing files
            foreach (var selectedItem in WebResourceItems.Where(w => w.Publish && w.SolutionId == solutionId))
            {
                selectedItems.Add(selectedItem);

                var item = ProjectFiles.FirstOrDefault(c => c.Content.ToString() == selectedItem.BoundFile);
                if (item == null) continue;

                var filePath = Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + selectedItem.BoundFile.Replace("/", "\\");
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"{Resource.MessageBox_CouldNotFindFile}: {selectedItem.BoundFile}");
                    return;
                }
            }

            UpdateWebResources(selectedItems);
        }

        private async void UpdateWebResources(List<WebResourceItem> items)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_UpdatingPublishing}...", vsStatusAnimation.vsStatusAnimationDeploy);

                var webResources = new List<Entity>();
                foreach (var webResourceItem in items)
                {
                    var webResource = WebResource.CreateUpdateWebResourceEntity(webResourceItem.WebResourceId,
                        webResourceItem.BoundFile, webResourceItem.Description, ConnPane.SelectedProject);
                    webResources.Add(webResource);
                }

                bool success;
                //Check if < CRM 2011 UR12 (ExecuteMultiple)
                var version = ConnPane.CrmService.ConnectedOrgVersion;
                if (version.Major == 5 && version.Revision < 3200)
                    success = await Task.Run(() => WebResource.UpdateAndPublishSingle(ConnPane.CrmService, webResources));
                else
                    success = await Task.Run(() => WebResource.UpdateAndPublishMultiple(ConnPane.CrmService, webResources));

                if (success) return;

                MessageBox.Show(Resource.ErrorMessage_ErrorUpdatingPublishingWebResources);
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

        private void ConnPane_OnSolutionProjectRemoved(object sender, SolutionProjectRemovedEventArgs e)
        {
            var project = e.Project;
            if (ConnPane.SelectedProject == project)
                ResetForm();
        }

        private void UpdateWebResourceItemsBoundFile(string oldValue, string newValue)
        {
            foreach (var webResourceItem in WebResourceItems
                .Where(w => w.BoundFile?.Equals(oldValue, StringComparison.InvariantCultureIgnoreCase) == true))
            {
                webResourceItem.BoundFile = newValue;
            }
        }

        private void UpdateMovedWebResourceItemsBoundFile(string newValue)
        {
            foreach (var movedWebResourceItem in _movedWebResourceItems)
            {
                movedWebResourceItem.WebResourceItem.BoundFile = newValue;
                movedWebResourceItem.WebResourceItem.Publish = movedWebResourceItem.Publish;
            }
        }

        private void UpdateProjectFilesAfterChange(string oldName, string newName)
        {
            var projectFile = ProjectFiles.FirstOrDefault(p => p.Content.ToString().Equals(oldName, StringComparison.InvariantCultureIgnoreCase));
            if (projectFile == null)
                return;

            if (string.IsNullOrEmpty(newName))
                ProjectFiles.Remove(projectFile);
            else
                projectFile.Content = newName;
        }

        private void UpdateWebResourceItemsBoundFilePath(string oldName, string newName)
        {
            foreach (var webResourceItem in WebResourceItems.Where(
                w => w.BoundFile?.StartsWith(oldName) == true))
            {
                webResourceItem.BoundFile = string.IsNullOrEmpty(newName)
                    ? null
                    : webResourceItem.BoundFile.Replace(oldName, newName);
            }
        }

        private void UpdateProjectFilesPathsAfterChange(string oldItemName, string newItemPath)
        {
            var projectFilesToRename = ProjectFiles.Where(p => p.Content.ToString().StartsWith(oldItemName)).ToList();
            foreach (var comboBoxItem in projectFilesToRename)
            {
                if (string.IsNullOrEmpty(newItemPath))
                    ProjectFiles.Remove(comboBoxItem);
                else
                    comboBoxItem.Content = comboBoxItem.Content.ToString().Replace(oldItemName, newItemPath);
            }
        }

        private void ConnPane_OnProjectItemRenamed(object sender, ProjectItemRenamedEventArgs e)
        {
            var projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;
            if (projectItem.Name == null) return;

            var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
            var oldName = e.OldName;
            var itemType = new Guid(projectItem.Kind);
            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                var newItemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                var oldItemName = newItemName.Replace(Path.GetFileName(projectItem.Name), oldName).Replace("//", "/");

                UpdateWebResourceItemsBoundFile(oldItemName, newItemName);

                UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = LocalPathToCrmPath(projectPath, projectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                var index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

                UpdateWebResourceItemsBoundFilePath(oldItemPath, newItemPath);

                UpdateProjectFilesPathsAfterChange(oldItemPath, newItemPath);
            }
        }

        private void ConnPane_OnProjectItemAdded(object sender, ProjectItemAddedEventArgs e)
        {
            //Web Application project does not execute this on an item being moved
            var projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;
            if (projectItem.Name == null) return;

            var itemType = new Guid(projectItem.Kind);
            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                if (Path.GetExtension(projectItem.Name).Equals(".exclude", StringComparison.CurrentCultureIgnoreCase))
                    return;

                var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
                var newItemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                if (ProjectFiles == null)
                    ProjectFiles = new ObservableCollection<ComboBoxItem>();

                ProjectFiles.Add(new ComboBoxItem { Content = newItemName });

                ProjectFiles = new ObservableCollection<ComboBoxItem>(ProjectFiles.OrderBy(p => p.Content.ToString()));
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
                ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject, ProjectType.WebResource);
        }

        private void ConnPane_OnProjectItemRemoved(object sender, ProjectItemRemovedEventArgs e)
        {
            //Web application projects do not execute this when an item being moved
            _movedWebResourceItems = new List<MovedWebResourceItem>();

            var projectItem = e.ProjectItem;
            if (projectItem.ContainingProject != ConnPane.SelectedProject) return;
            if (projectItem.Name == null) return;

            var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);

            var itemType = new Guid(projectItem.Kind);
            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                var itemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);
                var solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                                  ExtensionConstants.DefaultSolutionId;

                _movedWebResourceItems.AddRange(WebResourceItems
                    .Where(w => w.BoundFile == itemName && w.SolutionId == solutionId).Select(n => new MovedWebResourceItem
                    {
                        WebResourceItem = n,
                        Publish = n.Publish
                    }));

                if (!Path.GetExtension(projectItem.Name).Equals(".exclude", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (_movedWebResourceItems.Any())
                    {
                        if (!_movedWebResourceItems[0].WebResourceItem.Locked)
                            UpdateWebResourceItemsBoundFile(itemName, null);
                    }
                    else
                        UpdateWebResourceItemsBoundFile(itemName, null);
                }

                UpdateProjectFilesAfterChange(itemName, null);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var itemName = LocalPathToCrmPath(projectPath,
                    projectItem.FileNames[1].TrimEnd(Path.DirectorySeparatorChar));

                var index = itemName.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                ProjectFolders.Remove(itemName);

                UpdateWebResourceItemsBoundFilePath(itemName, null);

                UpdateProjectFilesPathsAfterChange(itemName, null);
            }
        }

        private void ConnPane_OnProjectItemMoved(object sender, ProjectItemMovedEventArgs e)
        {
            //Web application projects do not execute this when an item being moved
            var postMoveProjectItem = e.PostMoveProjectItem;

            var oldItemName = e.PreMoveName;
            var projectPath = Path.GetDirectoryName(postMoveProjectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            var itemType = new Guid(postMoveProjectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                var newItemName = LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1]);
                if (newItemName == null) return;

                if (Path.GetExtension(newItemName).Equals(".exclude", StringComparison.CurrentCultureIgnoreCase))
                    return;

                UpdateMovedWebResourceItemsBoundFile(newItemName);

                UpdateProjectFilesAfterChange(oldItemName, newItemName);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = LocalPathToCrmPath(projectPath, postMoveProjectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                var index = newItemPath.LastIndexOf(postMoveProjectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                UpdateWebResourceItemsBoundFilePath(oldItemName, newItemPath);

                UpdateProjectFilesPathsAfterChange(oldItemName, newItemPath);

                ProjectFolders = ProjectWorker.GetProjectFolders(ConnPane.SelectedProject, ProjectType.WebResource);
            }
        }

        private void ResetForm()
        {
            ResetCollections();

            SolutionList.IsEnabled = false;
            SolutionList.ItemsSource = null;
            WebResourceGrid.IsEnabled = false;
            Publish.IsEnabled = false;
            var publishAll =
                DataGridHelpers.FindVisualChildren<CheckBox>(WebResourceGrid)
                    .FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll != null)
                publishAll.IsChecked = false;

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
            var publishAll = (CheckBox)sender;
            var isChecked = publishAll.ReturnValue();

            UpdateAllPublishChecks(isChecked);
        }

        private void UpdateAllPublishChecks(bool publish)
        {
            foreach (var webResourceItem in WebResourceItems)
            {
                if (webResourceItem.AllowPublish)
                    webResourceItem.Publish = publish;
            }
        }

        private void BoundFile_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var grid = (Grid)sender;
            var textBlock = (TextBlock)grid.Children[0];

            var webResourceId = new Guid(textBlock.Tag.ToString());
            FileId.Content = webResourceId;
            var solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                              ExtensionConstants.DefaultSolutionId;

            var webResourceItem = WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId && w.SolutionId == solutionId);

            if (webResourceItem?.Locked == true)
                return;

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

            FilePopup = ControlHelper.ShowFilePopup(FilePopup, grid);
            ProjectFileList = ControlHelper.ShowProjectFileList(ProjectFileList, WebResourceGrid.Columns[6].ActualWidth);
        }

        private async void GetWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            var webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
            var webResourceItem =
                WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);

            var folder = string.Empty;
            if (!string.IsNullOrEmpty(webResourceItem?.BoundFile))
                folder = WebResource.GetExistingFolderFromBoundFile(webResourceItem, folder);

            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService);
        }

        private async void DownloadWebResourceToFolder(string folder, Guid webResourceId)
        {
            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService);

            ProjectFolderList.SelectedItem = null;
        }

        private void DownloadAll_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_DownloadingFiles}...", vsStatusAnimation.vsStatusAnimationSync);

                var result = MessageBox.Show(Resource.Message_DownloadAllWebResources +
                                                          Environment.NewLine + Environment.NewLine + Resource.Message_OkProceed, Resource.MessageBox_WebResourceDownload,
                    MessageBoxButton.YesNo);

                if (result != MessageBoxResult.Yes)
                    return;

                var icv = CollectionViewSource.GetDefaultView(WebResourceGrid.ItemsSource);
                var downloadItems = icv.Cast<WebResourceItem>().ToList();

                var updatesToMake = new Dictionary<Guid, string>();

                Parallel.ForEach(downloadItems, currentItem =>
                {
                    var webResource = WebResource.RetrieveWebResourceFromCrm(ConnPane.CrmService, currentItem.WebResourceId);

                    var name = webResource.GetAttributeValue<string>("name");
                    name = WebResource.AddMissingExtension(name,
                        webResource.GetAttributeValue<OptionSetValue>("webresourcetype").Value);

                    //TODO: option to change root folder
                    var path = WebResource.ConvertWebResourceNameFullToPath(name, "/", ConnPane.SelectedProject);

                    var decodedContent = WebResource.GetDecodedContent(webResource);
                    WriteFileToDisk(path, decodedContent);

                    var projectItem = ConnPane.SelectedProject.ProjectItems.AddFromFile(path);

                    var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
                    var fullname = projectItem.FileNames[1];
                    var boundName = fullname.Replace(projectPath, string.Empty).Replace("\\", "/");

                    updatesToMake.Add(currentItem.WebResourceId, boundName);
                });

                foreach (var update in updatesToMake)
                {
                    foreach (var item in WebResourceItems.Where(w => w.WebResourceId == update.Key))
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
                Overlay.ShowMessage(_dte, $"{Resource.Message_DownloadingFile}...", vsStatusAnimation.vsStatusAnimationSync);

                var webResource = await Task.Run(() => WebResource.RetrieveWebResourceFromCrm(client, webResourceId));

                var name = webResource.GetAttributeValue<string>("name");
                name = WebResource.AddMissingExtension(name, webResource.GetAttributeValue<OptionSetValue>("webresourcetype").Value);

                var path = WebResource.ConvertWebResourceNameToPath(name, folder, ConnPane.SelectedProject.FullName);

                if (File.Exists(path))
                {
                    var result = MessageBox.Show(Resource.MessageBox_OkOverwrite, Resource.MessageBox_WebResourceDownload,
                        MessageBoxButton.YesNo);
                    if (result != MessageBoxResult.Yes)
                        return;
                }

                var decodedContent = WebResource.GetDecodedContent(webResource);
                WriteFileToDisk(path, decodedContent);

                var projectItem = ConnPane.SelectedProject.ProjectItems.AddFromFile(path);

                var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
                var fullname = projectItem.FileNames[1];
                var boundName = fullname.Replace(projectPath, string.Empty).Replace("\\", "/");
                if (!boundName.StartsWith("/"))
                    boundName = $"/{boundName}";

                foreach (var item in WebResourceItems.Where(w => w.WebResourceId == webResourceId))
                {
                    item.BoundFile = boundName;

                    var publishAll =
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
            var webResourceId = new Guid(((Button)sender).CommandParameter.ToString());

            var contentUrl = $"main.aspx?etc=9333&id=%7b{webResourceId}%7d&pagetype=webresourceedit";

            WebBrowser.OpenCrmPage(ConnPane.CrmService, contentUrl);
        }

        private async void CompareWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_DownloadingFileForCompare}...", vsStatusAnimation.vsStatusAnimationSync);

                var webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
                var webResource = await Task.Run(() => WebResource.RetrieveWebResourceContentFromCrm(ConnPane.CrmService, webResourceId));

                var tempFile = WriteTempFile(webResource.GetAttributeValue<string>("name"),
                        WebResource.DecodeWebResource(webResource.GetAttributeValue<string>("content")));

                var webResourceItem = WebResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);
                if (webResourceItem == null)
                    return;

                var boundFilePath = webResourceItem.BoundFile;

                var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
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
            var deleteResult = MessageBox.Show(Resource.Message_AreYouSure + Environment.NewLine + Environment.NewLine +
                                                            Resource.Message_WillAttemptToDelete, Resource.Message_DeleteWebResource, MessageBoxButton.YesNo);
            if (deleteResult != MessageBoxResult.Yes)
                return;

            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_DeletingWebResource}...", vsStatusAnimation.vsStatusAnimationSync);

                var webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
                await Task.Run(() => WebResource.DeleteWebResourceFromCrm(ConnPane.CrmService, webResourceId));

                WebResourceItems = new ObservableCollection<WebResourceItem>(WebResourceItems.Where(w => w.WebResourceId != webResourceId));

                WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

                FilterWebResourceItems();
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void ProjectFileList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectFileList.SelectedIndex == -1)
            {
                FilePopup.IsOpen = false;
                return;
            }

            var webResourceItem =
                WebResourceItems.FirstOrDefault(w => w.WebResourceId == new Guid(FileId.Content.ToString()));

            var item = (ComboBoxItem)ProjectFileList.SelectedItem;

            if (webResourceItem != null && webResourceItem.BoundFile != item.Content.ToString())
                webResourceItem.BoundFile = item.Content.ToString();

            FilePopup.IsOpen = false;
        }

        private async Task GetCrmData()
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_RetrievingSolutionsWebResources}...",
                    vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();
                var webResourceTask = GetWebResources();

                await Task.WhenAll(solutionTask, webResourceTask);

                if (!solutionTask.Result)
                {
                    MessageBox.Show(Resource.ErrorMessage_ErrorRetrievingSolutions);
                    return;
                }

                if (!webResourceTask.Result)
                    MessageBox.Show(Resource.ErrorMessage_ErrorRetrievingWebResources);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task<bool> GetSolutions()
        {
            var results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService, true));
            if (results == null)
                return false;

            var solutions = ModelBuilder.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = solutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private async Task<bool> GetWebResources()
        {
            var results = await Task.Run(() => WebResource.RetrieveWebResourcesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            WebResourceItems = ModelBuilder.CreateWebResourceItemView(results, ConnPane.SelectedProject.Name);

            CreateFilterTypeNameList();
            CreateFilterStateList();

            WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

            foreach (var webResourceItem in WebResourceItems)
            {
                webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
            }

            WebResourceGrid.IsEnabled = true;
            FilterWebResourceItems();

            return true;
        }

        private void WebResourceItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var item = (WebResourceItem)sender;
            var solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                              ExtensionConstants.DefaultSolutionId;

            if (e.PropertyName == "BoundFile")
            {
                if (WebResourceGrid.ItemsSource != null)
                {
                    foreach (var webResourceItem in WebResourceItems.Where(w =>
                        w.Name == item.Name && w.BoundFile != item.BoundFile))
                    {
                        webResourceItem.PropertyChanged -= WebResourceItem_PropertyChanged;
                        webResourceItem.BoundFile = item.BoundFile;
                        webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
                        if (string.IsNullOrEmpty(item.BoundFile) && item.Publish)
                            webResourceItem.Publish = false;
                    }
                }

                Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, item);
            }

            if (e.PropertyName == "Publish")
            {
                foreach (var webResourceItem in WebResourceItems.Where(w => w.WebResourceId == item.WebResourceId && w.SolutionId == solutionId))
                {
                    webResourceItem.Publish = item.Publish;
                }

                Publish.IsEnabled = WebResourceItems.Count(w => w.Publish) > 0;
                ControlHelper.SetPublishAll(WebResourceGrid);
            }

            if (e.PropertyName == "Locked")
            {
                foreach (var webResourceItem in WebResourceItems.Where(w => w.WebResourceId == item.WebResourceId && w.SolutionId == solutionId))
                {
                    webResourceItem.Locked = item.Locked;
                }
            }
        }

        private void CreateFilterTypeNameList()
        {
            FilterTypeNames = FilterTypeName.CreateFilterList(ConnPane.CrmService.ConnectedOrgVersion.Major);

            foreach (var filterTypeName in FilterTypeNames)
            {
                filterTypeName.PropertyChanged += Filter_PropertyChanged;
            }
        }

        private void CreateFilterStateList()
        {
            FilterStates = FilterState.CreateFilterList();

            foreach (var filterTypeState in FilterStates)
            {
                filterTypeState.PropertyChanged += Filter_PropertyChanged;
            }
        }

        private void Filter_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var type = sender.GetType();

            if (type == typeof(FilterState))
                GridFilters.SetSelectAll(sender, FilterStates);
            else if (type == typeof(FilterTypeName))
                GridFilters.SetSelectAll(sender, FilterTypeNames);

            FilterWebResourceItems();
        }

        public void FilterWebResourceItems()
        {
            if (WebResourceItems.Count == 0)
                return;

            var icv = CollectionViewSource.GetDefaultView(WebResourceGrid.ItemsSource);
            if (icv == null) return;

            icv.Filter = GetFilteredView;

            //Only keep publish flags for items still visible
            foreach (var item in WebResourceItems.Where(w => w.Publish).Except(icv.OfType<WebResourceItem>()))
            {
                item.Publish = false;
            }
        }

        public bool GetFilteredView(object sourceObject)
        {
            var filterCriteria = new FilterCriteria
            {
                WebResourceItem = (WebResourceItem)sourceObject,
                FilterTypeNames = FilterTypeNames,
                FilterStates = FilterStates,
                SearchText = Search.Text.Trim(),
                SolutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                             ExtensionConstants.DefaultSolutionId
            };

            return DataFilter.FilterItems(filterCriteria);
        }

        private void Refresh_OnClick(object sender, RoutedEventArgs e)
        {
            RefreshFromCrm();
        }

        private async void RefreshFromCrm()
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_RetrievingWebResources}...", vsStatusAnimation.vsStatusAnimationSync);

                var toPublish = WebResourceItems.Where(w => w.Publish);

                var results = await Task.Run(() => WebResource.RetrieveWebResourcesFromCrm(ConnPane.CrmService));
                if (results == null)
                {
                    MessageBox.Show(Resource.ErrorMessage_ErrorRetrievingWebResources);
                    return;
                }

                WebResourceItems = ModelBuilder.CreateWebResourceItemView(results, ConnPane.SelectedProject.Name);

                foreach (var webResourceItem in WebResourceItems)
                {
                    webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
                }

                WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);

                FilterWebResourceItems();

                WebResourceItems = WebResourceItemHandler.ResetPublishValues(toPublish, WebResourceItems);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void ProjectFolderList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var projectFolderList = (ComboBox)sender;
            var folder = projectFolderList.SelectedValue?.ToString();

            if (string.IsNullOrEmpty(folder))
            {
                FolderPopup.IsOpen = false;
                return;
            }

            var webResourceId = new Guid(FolderId.Content.ToString());
            DownloadWebResourceToFolder(folder, webResourceId);
            FolderPopup.IsOpen = false;
        }

        private void GetWebResource_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var button = (Button)sender;
            FolderId.Content = button.CommandParameter;
            FolderPopup = ControlHelper.ShowFolderPopup(FolderPopup, button);
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
            var editDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "EditDescription");
            editDescription.Visibility = editVisibility;

            var cancelEditDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "CancelEditDescription");
            cancelEditDescription.Visibility = saveCancelVisibility;

            var saveDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "SaveDescription");
            saveDescription.Visibility = saveCancelVisibility;

            var undoEditDescription = DetailsRow.GetDataGridRowControl<Button>(sender, "UndoEditDescription");
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

            var webResourceItem = WebResourceItemHandler.WebResourceItemFromCmdParam(sender, WebResourceItems);
            WebResourceItems = WebResourceItemHandler.ResetDescriptions(WebResourceItems, webResourceItem);
        }

        private void SaveDescription_Click(object sender, RoutedEventArgs e)
        {
            ShowHideDescriptionButtons(sender, Visibility.Visible, Visibility.Hidden);
            DetailsRow.ShowHideDetailsRow(sender);

            var webResourceItem = WebResourceItemHandler.WebResourceItemFromCmdParam(sender, WebResourceItems);

            var description = DetailsRow.GetDataGridRowControl<TextBox>(sender, "Description");
            var newDescription = description.Text.Trim();
            webResourceItem = WebResourceItemHandler.SetDescriptionFromInput(webResourceItem, newDescription);
            WebResourceItems = WebResourceItemHandler.SetDescriptions(WebResourceItems, webResourceItem.WebResourceId, newDescription);

            var webResources = WebResource.CreateDescriptionUpdateWebResource(webResourceItem, newDescription);

            ExecuteSaveDescription(webResources, webResourceItem);
        }

        private async void ExecuteSaveDescription(List<Entity> webResources, WebResourceItem webResourceItem)
        {
            await Task.Run(() => SaveDescription(webResources));

            Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, webResourceItem);
        }

        private void SaveDescription(List<Entity> webResources)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_Updating}...", vsStatusAnimation.vsStatusAnimationDeploy);

                WebResource.UpdateAndPublishSingle(ConnPane.CrmService, webResources);
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

            var webResourceItem = WebResourceItemHandler.WebResourceItemFromCmdParam(sender, WebResourceItems);

            ExecuteUndoDescription(webResourceItem);
        }

        private async void ExecuteUndoDescription(WebResourceItem webResourceItem)
        {
            var serverDescription = await Task.Run(() => UndoDescription(webResourceItem));

            WebResourceItems = WebResourceItemHandler.SetDescriptions(WebResourceItems, webResourceItem.WebResourceId, serverDescription);

            Mapping.AddOrUpdateSpklMapping(ConnPane.SelectedProject, ConnPane.SelectedProfile, webResourceItem);
        }

        private string UndoDescription(WebResourceItem webResourceItem)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_Updating}...", vsStatusAnimation.vsStatusAnimationDeploy);

                var serverDescription =
                    WebResource.RetrieveWebResourceDescriptionFromCrm(ConnPane.CrmService, webResourceItem.WebResourceId);

                return serverDescription;
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void ShowId_Click(object sender, RoutedEventArgs e)
        {
            var width = WebResourceGrid.Columns[0].Width;

            WebResourceGrid.Columns[0].Width = width.UnitType == DataGridLengthUnitType.SizeToCells
                ? new DataGridLength(25, DataGridLengthUnitType.Pixel)
                : DataGridLength.SizeToCells;

            ControlHelper.RotateButtonImage((Button)sender, width);
        }

        private void ConnPane_ProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ConnPane.CrmService?.IsReady != true)
                return;

            foreach (var webResourceItem in WebResourceItems
                .Where(w => w.Publish || w.BoundFile != null || w.Description != null
                || w.PreviousDescription != null))
            {
                webResourceItem.PropertyChanged -= WebResourceItem_PropertyChanged;
                webResourceItem.Publish = false;
                webResourceItem.BoundFile = null;
                webResourceItem.Description = null;
                webResourceItem.PreviousDescription = null;
                webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;
            }

            WebResourceItems = Mapping.HandleSpklMappings(ConnPane.SelectedProject, ConnPane.SelectedProfile, WebResourceItems);
        }

        private void TextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }

        private void ClearFilters_Click(object sender, RoutedEventArgs e)
        {
            FilterStates = FilterState.ResetFilter(FilterStates);
            FilterTypeNames = FilterTypeName.ResetFilter(FilterTypeNames);
            Search.Text = string.Empty;
        }
    }
}