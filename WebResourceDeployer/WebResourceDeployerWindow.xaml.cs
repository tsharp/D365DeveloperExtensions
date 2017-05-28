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
using System.Windows.Threading;
using CrmDeveloperExtensions.Core.Connection;
using CrmDeveloperExtensions.Core.Enums;
using CrmDeveloperExtensions.Core.Logging;
using CrmDeveloperExtensions.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using Microsoft.VisualStudio;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using WebResourceDeployer.ViewModels;
using static CrmDeveloperExtensions.Core.FileSystem;
using Task = System.Threading.Tasks.Task;

namespace WebResourceDeployer
{
    /// <summary>
    /// Interaction logic for WebResourceDeployerWindow.xaml
    /// </summary>
    public partial class WebResourceDeployerWindow : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly IVsSolution _vsSolution;
        private List<WebResourceItem> _movedWebResourceItems;
        private List<string> _movedBoundFiles;
        private uint _movedItemid;
        private bool _projectEventsRegistered;
        private readonly Logger _logger;
        private readonly FieldInfo _menuDropAlignmentField;
        private const string WindowType = "WebResourceDeployer";
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();

        private ObservableCollection<WebResourceItem> _items;
        private ObservableCollection<ComboBoxItem> _projectFiles;

        public ObservableCollection<ComboBoxItem> ProjectFiles
        {
            get => _projectFiles;
            set
            {
                _projectFiles = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
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

            //Fix for Tablet/Touchscreen left-right menu
            _menuDropAlignmentField = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static);
            System.Diagnostics.Debug.Assert(_menuDropAlignmentField != null);
            EnsureStandardPopupAlignment();
            SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
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

            //if (!_projectEventsRegistered)
            //{
            //    RegisterProjectEvents();
            //    _projectEventsRegistered = true;
            //}
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            _items = new ObservableCollection<WebResourceItem>();
            ProjectFiles = new ObservableCollection<ComboBoxItem>();
            ProjectFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);

            await GetCrmData();

            SolutionList.IsEnabled = true;

            if (!CrmDeveloperExtensions.Core.Config.ConfigFile.ConfigFileExists(_dte.Solution.FullName))
                CrmDeveloperExtensions.Core.Config.ConfigFile.CreateConfigFile(ConnPane.OrganizationId, ConnPane.SelectedProject.UniqueName, _dte.Solution.FullName);
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FilterWebResources();
        }

        private void WebResourceType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FilterWebResources();
        }

        private void AddWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            ObservableCollection<ComboBoxItem> projectsFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;
            NewWebResource newWebResource = new NewWebResource(ConnPane.CrmService, projectsFiles, solutionId);
            bool? result = newWebResource.ShowModal();

            if (result != true) return;

            ObservableCollection<MenuItem> projectFolders = GetProjectFolders();
            WebResourceItem defaultItem = Class1.WebResourceItemFromNew(newWebResource, newWebResource.NewSolutionId, projectFolders);
            defaultItem.PropertyChanged += WebResourceItem_PropertyChanged;
            //Needs to be after setting the property changed event
            defaultItem.BoundFile = newWebResource.NewBoundFile;

            foreach (MenuItem menuItem in defaultItem.ProjectFolders)
                menuItem.CommandParameter = defaultItem.WebResourceId;

            _items.Add(defaultItem);

            if (newWebResource.NewSolutionId != CrmDeveloperExtensions.Core.ExtensionConstants.DefaultSolutionId)
            {
                WebResourceItem solutionItem = Class1.WebResourceItemFromNew(newWebResource, CrmDeveloperExtensions.Core.ExtensionConstants.DefaultSolutionId, projectFolders);
                solutionItem.PropertyChanged += WebResourceItem_PropertyChanged;
                //Needs to be after setting the property changed event
                solutionItem.BoundFile = newWebResource.NewBoundFile;

                foreach (MenuItem menuItem in solutionItem.ProjectFolders)
                    menuItem.CommandParameter = solutionItem.WebResourceId;

                _items.Add(solutionItem);
            }

            WebResourceGrid.ItemsSource = _items.OrderBy(w => w.Name).ToList();

            var filter = WebResourceType.SelectedValue;
            var showManaged = ShowManaged.IsChecked;

            FilterWebResources();

            WebResourceType.SelectedValue = filter;
            ShowManaged.IsChecked = showManaged;

            WebResourceGrid.ScrollIntoView(defaultItem);
        }

        private void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            List<WebResourceItem> selectedItems = new List<WebResourceItem>();
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ?? Guid.Empty;

            //Check for unsaved & missing files
            List<ProjectItem> dirtyItems = new List<ProjectItem>();
            foreach (var selectedItem in _items.Where(w => w.Publish && w.SolutionId == solutionId))
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
            ShowMessage("Updating & Publishing...", vsStatusAnimation.vsStatusAnimationDeploy);

            List<Entity> webResources = new List<Entity>();
            foreach (WebResourceItem webResourceItem in items)
            {
                Entity webResource = Crm.WebResource.CreateUpdateWebResourceEntity(webResourceItem.WebResourceId,
                    webResourceItem.BoundFile, ConnPane.SelectedProject.FullName);
                webResources.Add(webResource);
            }

            bool success;
            //Check if < CRM 2011 UR12 (ExecuteMutliple)
            Version version = ConnPane.CrmService.ConnectedOrgVersion;
            if (version.Major == 5 && version.Revision < 3200)
                success = await Task.Run(() => Crm.WebResource.UpdateAndPublishSingle(ConnPane.CrmService, webResources));
            else
                success = await Task.Run(() => Crm.WebResource.UpdateAndPublishMultiple(ConnPane.CrmService, webResources));

            HideMessage(vsStatusAnimation.vsStatusAnimationDeploy);

            if (success) return;

            MessageBox.Show("Error Updating And Publishing Web Resources. See the Output Window for additional details.");
        }

        private void ShowManaged_OnChecked(object sender, RoutedEventArgs e)
        {
            FilterWebResources();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            CrmDeveloperExtensions.Core.WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService.CrmConnectOrgUriActual,
                $"tools/solution/edit.aspx?id=%7b{CrmDeveloperExtensions.Core.ExtensionConstants.DefaultSolutionId}%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            CrmDeveloperExtensions.Core.WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService.CrmConnectOrgUriActual,
                "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void WebResourceGrid_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Make rows unselectable
            WebResourceGrid.UnselectAllCells();
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();
        }

        private void ConnPane_OnSolutionProjectRemoved(object sender, SolutionProjectRemovedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void ConnPane_OnProjectItemRenamed(object sender, ProjectItemRenamedEventArgs e)
        {
            ProjectItem projectItem = e.ProjectItem;
            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            string oldName = e.OldName;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType == VSConstants.GUID_ItemType_PhysicalFile)
            {
                string newItemName = LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

                if (projectItem.Name == null) return;

                var oldItemName = newItemName.Replace(Path.GetFileName(projectItem.Name), oldName).Replace("//", "/");

                foreach (WebResourceItem webResourceItem in _items.Where(w => w.BoundFile != null && w.BoundFile.Equals(oldItemName, StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (webResourceItem.BoundFile != oldItemName) continue;

                    webResourceItem.BoundFile = newItemName;
                }

                ProjectFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);
            }

            if (itemType == VSConstants.GUID_ItemType_PhysicalFolder)
            {
                var newItemPath = LocalPathToCrmPath(projectPath, projectItem.FileNames[1])
                    .TrimEnd(Path.DirectorySeparatorChar);

                int index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
                if (index == -1) return;

                var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

                foreach (WebResourceItem webResourceItem in _items.Where(w => w.BoundFile != null && w.BoundFile.StartsWith(oldItemPath)))
                    webResourceItem.BoundFile = webResourceItem.BoundFile.Replace(oldItemPath, newItemPath);

                ProjectFiles = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);

            }
        }

        private void ConnPane_OnProjectItemAdded(object sender, ProjectItemAddedEventArgs e)
        {
            //Add to web resource item - project file list
        }

        private void ConnPane_OnProjectItemRemoved(object sender, ProjectItemRemovedEventArgs e)
        {
            //Remove mappings
            //Remove from web resource item - project file list
        }

        private void ResetForm()
        {
            _items = new ObservableCollection<WebResourceItem>();
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            SolutionList.IsEnabled = false;
            AddWebResource.IsEnabled = false;
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
            foreach (WebResourceItem webResourceItem in _items)
                if (webResourceItem.AllowPublish)
                    webResourceItem.Publish = publish;
        }

        private void BoundFile_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Grid grid = (Grid)sender;
            TextBlock textBlock = (TextBlock)grid.Children[0];

            Guid webResourceId = new Guid(textBlock.Tag.ToString());
            FileId.Content = webResourceId;

            WebResourceItem webResourceItem = _items.FirstOrDefault(w => w.WebResourceId == webResourceId);
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

            ProjectFileList.Width = WebResourceGrid.Columns[5].ActualWidth - 2;
            FilePopup.PlacementTarget = textBlock;
            FilePopup.Placement = PlacementMode.Relative;
            FilePopup.IsOpen = true;
            ProjectFileList.IsDropDownOpen = true;
        }

        private async void GetWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
            WebResourceItem webResourceItem =
                _items.FirstOrDefault(w => w.WebResourceId == webResourceId);

            string folder = String.Empty;
            if (!string.IsNullOrEmpty(webResourceItem?.BoundFile))
                folder = Crm.WebResource.GetExistingFolderFromBoundFile(webResourceItem, folder);

            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService, ConnPane.SelectedProject.Name);
        }

        private async void DownloadWebResourceToFolder(object sender, RoutedEventArgs routedEventArgs)
        {
            MenuItem item = (MenuItem)sender;
            string folder = item.Header.ToString();
            Guid webResourceId = (Guid)item.CommandParameter;

            await DownloadWebResourceAsync(webResourceId, folder, ConnPane.CrmService, ConnPane.SelectedProject.Name);
        }

        private async Task DownloadWebResourceAsync(Guid webResourceId, string folder, CrmServiceClient client, string projectName)
        {
            try
            {
                ShowMessage("Downloading file...", vsStatusAnimation.vsStatusAnimationSync);

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
                        HideMessage(vsStatusAnimation.vsStatusAnimationSync);
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

                foreach (WebResourceItem item in _items.Where(w => w.WebResourceId == webResourceId))
                {
                    item.BoundFile = boundName;
                    item.AllowCompare = SetAllowCompare(item.Type);

                    CheckBox publishAll =
                        FindVisualChildren<CheckBox>(WebResourceGrid)
                            .FirstOrDefault(t => t.Name == "PublishSelectAll");
                    if (publishAll == null) return;

                    if (publishAll.IsChecked == true)
                        item.Publish = true;
                }
            }
            finally
            {
                HideMessage(vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void OpenWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());

            Uri crmFullUri = ConnPane.CrmService.CrmConnectOrgUriActual;
            Uri crmUrl = new Uri(crmFullUri.GetLeftPart(UriPartial.Authority));
            string contentUrl = $"main.aspx?etc=9333&id=%7b{webResourceId}%7d&pagetype=webresourceedit";

            CrmDeveloperExtensions.Core.WebBrowser.OpenCrmPage(_dte, crmUrl, contentUrl);
        }

        private async void CompareWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            ShowMessage("Downloading file for compare...", vsStatusAnimation.vsStatusAnimationSync);

            //Get the file from CRM and save in temp files
            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
            Entity webResource = await Task.Run(() => Crm.WebResource.RetrieveWebResourceContentFromCrm(ConnPane.CrmService, webResourceId));

            OutputLogger.WriteToOutputWindow($"Retrieved Web Resource {webResourceId} For Compare", MessageType.Info);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);

            string tempFile = WriteTempFile(webResource.GetAttributeValue<string>("name"),
                    Crm.WebResource.DecodeWebResource(webResource.GetAttributeValue<string>("content")));

            //string projectName = ConnPane.SelectedProject.Name;
            //Project project = GetProjectByName(projectName);
            //var projectPath = Path.GetDirectoryName(project.FullName);
            //if (projectPath == null) return;

            //string boundFilePath = String.Empty;
            //List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            //foreach (WebResourceItem webResourceItem in webResources)
            //{
            //    if (webResourceItem.WebResourceId == webResourceId)
            //        boundFilePath = webResourceItem.BoundFile;
            //}

            //_dte.ExecuteCommand("Tools.DiffFiles",
            //    string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\"", tempFile,
            //        projectPath + boundFilePath.Replace("/", "\\"),
            //        webResource.GetAttributeValue<string>("name") + " - CRM", boundFilePath + " - Local"));
        }

        private async void DeleteWebResource_OnClick(object sender, RoutedEventArgs e)
        {
            MessageBoxResult deleteResult = MessageBox.Show("Are you sure?" + Environment.NewLine + Environment.NewLine +
                                                            "This will attempt to delete the web resource from CRM.", "Delete Web Resource", MessageBoxButton.YesNo);
            if (deleteResult != MessageBoxResult.Yes) return;

            ShowMessage("Deleting web resource...", vsStatusAnimation.vsStatusAnimationSync);

            Guid webResourceId = new Guid(((Button)sender).CommandParameter.ToString());
            await Task.Run(() => Crm.WebResource.DeleteWebResourcetFromCrm(ConnPane.CrmService, webResourceId));

            //List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            //if (webResources == null) return;

            foreach (WebResourceItem webResourceItem in _items.Where(w => w.WebResourceId == webResourceId))
                _items.Remove(webResourceItem);
            ////webResources = HandleMappings(webResources);
            //WebResourceGrid.ItemsSource = webResources;

            FilterWebResources();

            OutputLogger.WriteToOutputWindow($"Deleted Web Resource {webResourceId}", MessageType.Info);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);
        }

        private void ProjectFileList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectFileList.SelectedIndex == -1) return;

            WebResourceItem webResourceItem =
                _items.FirstOrDefault(w => w.WebResourceId == new Guid(FileId.Content.ToString()));

            ComboBoxItem item = (ComboBoxItem)ProjectFileList.SelectedItem;

            if (webResourceItem != null && webResourceItem.BoundFile != item.Content.ToString())
                webResourceItem.BoundFile = item.Content.ToString();

            FilePopup.IsOpen = false;
        }

        private async Task GetCrmData()
        {
            ShowMessage("Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

            var solutionTask = GetSolutions();
            var webResourceTask = GetWebResources();

            await Task.WhenAll(solutionTask, webResourceTask);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);

            if (!solutionTask.Result)
            {
                MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                return;
            }

            if (!webResourceTask.Result)
                MessageBox.Show("Error Retrieving Web Resources. See the Output Window for additional details.");
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService, true));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            List<CrmSolution> solutions = Class1.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = solutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private ObservableCollection<MenuItem> GetProjectFolders()
        {
            ObservableCollection<MenuItem> projectFolders = ProjectWorker.GetProjectFoldersForMenu(ConnPane.SelectedProject.Name);
            foreach (MenuItem projectFolder in projectFolders)
            {
                projectFolder.Click += DownloadWebResourceToFolder;
            }

            return projectFolders;
        }

        private async Task<bool> GetWebResources()
        {
            EntityCollection results = await Task.Run(() => Crm.WebResource.RetrieveWebResourcesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Web Resources From CRM", MessageType.Info);

            ObservableCollection<MenuItem> projectFolders = GetProjectFolders();

            _items = Class1.CreateWebResourceItemView2(results, ConnPane.SelectedProject.Name, projectFolders);

            foreach (WebResourceItem webResourceItem in _items)
                webResourceItem.PropertyChanged += WebResourceItem_PropertyChanged;

            _items = new ObservableCollection<WebResourceItem>(_items.OrderBy(w => w.Name));

            _items = Mapping.HandleMappings(_dte, ConnPane.SelectedProject, _items, ConnPane.OrganizationId);
            WebResourceGrid.ItemsSource = _items;
            FilterWebResources();
            WebResourceGrid.IsEnabled = true;
            WebResourceType.IsEnabled = true;
            ShowManaged.IsEnabled = true;
            AddWebResource.IsEnabled = true;

            return true;
        }

        private void WebResourceItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            WebResourceItem item = (WebResourceItem)sender;

            if (e.PropertyName == "BoundFile")
            {
                if (WebResourceGrid.ItemsSource != null)
                {
                    foreach (WebResourceItem webResourceItem in _items.Where(w => w.WebResourceId == item.WebResourceId))
                        webResourceItem.BoundFile = item.BoundFile;
                }

                Mapping.AddOrUpdateMapping(_dte.Solution.FullName, ConnPane.SelectedProject, item, ConnPane.OrganizationId);
            }

            if (e.PropertyName == "Publish")
            {
                foreach (WebResourceItem webResourceItem in _items.Where(w => w.WebResourceId == item.WebResourceId))
                    webResourceItem.Publish = item.Publish;

                Publish.IsEnabled = _items.Count(w => w.Publish) > 0;

                SetPublishAll();
            }
        }

        private void SetPublishAll()
        {
            //List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            //if (webResources == null) return;

            //Set Publish All
            CheckBox publishAll = FindVisualChildren<CheckBox>(WebResourceGrid).FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll == null) return;

            publishAll.IsChecked = _items.Count(w => w.Publish) == _items.Count(w => w.AllowPublish);
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                if (child != null && child is T)
                    yield return (T)child;

                foreach (T childOfChild in FindVisualChildren<T>(child))
                    yield return childOfChild;
            }
        }

        private void FilterWebResources()
        {
            ComboBoxItem selectedItem = (ComboBoxItem)WebResourceType.SelectedItem;
            string type = selectedItem?.Tag.ToString() ?? String.Empty;
            bool showManaged = ShowManaged.IsChecked != null && ShowManaged.IsChecked.Value;
            Guid solutionId = ((CrmSolution)SolutionList.SelectedItem)?.SolutionId ??
                CrmDeveloperExtensions.Core.ExtensionConstants.DefaultSolutionId;

            //Clear publish flags
            if (!string.IsNullOrEmpty(type))
            {
                foreach (WebResourceItem webResourceItem in _items)
                {
                    if (selectedItem != null && (webResourceItem.Type.ToString() != selectedItem.Tag.ToString() || (webResourceItem.IsManaged && !showManaged)))
                        webResourceItem.Publish = false;
                }
            }

            //Filter the items
            ICollectionView icv = CollectionViewSource.GetDefaultView(WebResourceGrid.ItemsSource);
            if (icv == null) return;

            icv.Filter = o =>
            {
                WebResourceItem w = o as WebResourceItem;
                //File type filter & show managed + unmanaged
                if (!string.IsNullOrEmpty(type) && showManaged)
                    return w != null && (w.Type.ToString() == type && w.SolutionId == solutionId);

                //File type filter & show unmanaged only
                if (!string.IsNullOrEmpty(type) && !showManaged)
                    return w != null && (w.Type.ToString() == type && !w.IsManaged && w.SolutionId == solutionId);

                //No file type filter & show managed + unmanaged
                if (string.IsNullOrEmpty(type) && showManaged)
                    return w != null && (w.SolutionId == solutionId);

                //No file type filter & show unmanaged only
                return w != null && (!w.IsManaged && w.SolutionId == solutionId);
            };

            //Item Count
            CollectionView cv = (CollectionView)icv;
            ItemCount.Text = cv.Count + " Items";
        }

        private void ShowMessage(string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            CrmDeveloperExtensions.Core.StatusBar.SetStatusBarValue(_dte, "Retrieving web resources...", (vsStatusAnimation)animation);
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
                            CrmDeveloperExtensions.Core.StatusBar.ClearStatusBarValue(_dte, (vsStatusAnimation)animation);
                        LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }

        private static bool SetAllowCompare(int type)
        {
            int[] noCompare = { 5, 6, 7, 8, 10 };
            return !noCompare.Contains(type);
        }
    }
}
