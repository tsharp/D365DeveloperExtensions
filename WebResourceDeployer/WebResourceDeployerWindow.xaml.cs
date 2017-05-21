using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using CrmDeveloperExtensions.Core.Connection;
using CrmDeveloperExtensions.Core.Enums;
using CrmDeveloperExtensions.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using EnvDTE80;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using WebResourceDeployer.ViewModels;
using Task = System.Threading.Tasks.Task;

namespace WebResourceDeployer
{
    /// <summary>
    /// Interaction logic for WebResourceDeployerWindow.xaml
    /// </summary>
    public partial class WebResourceDeployerWindow : UserControl
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

        public WebResourceDeployerWindow()
        {
            InitializeComponent();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            //var events = _dte.Events;
            //var windowEvents = events.WindowEvents;
            //windowEvents.WindowActivated += WindowEventsOnWindowActivated;
            //var solutionEvents = events.SolutionEvents;
            //solutionEvents.BeforeClosing += SolutionBeforeClosing;
            //solutionEvents.SolutionProjectRemoved += SolutionProjectRemoved;

            //var events2 = (Events2)_dte.Events;
            //var projectItemsEvents = events2.ProjectItemsEvents;
            //projectItemsEvents.ItemRenamed += ProjectItemRenamed;

            //IVsSolutionEvents vsSolutionEvents = new VsSolutionEvents(this);
            //_vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            //_vsSolution.AdviseSolutionEvents(vsSolutionEvents, out solutionEventsCookie);

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
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            //await GetWebResources();
            await GetCrmData();

            ProjectFileList.ItemsSource = ProjectWorker.GetProjectFilesForComboBox(ConnPane.SelectedProject);

            SolutionList.IsEnabled = true;
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
            throw new NotImplementedException();
        }

        private void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void ShowManaged_OnChecked(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            CrmDeveloperExtensions.Core.WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService.CrmConnectOrgUriActual,
                "tools/solution/edit.aspx?id=%7bfd140aaf-4df4-11dd-bd17-0019b9312238%7d");
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
            throw new NotImplementedException();
        }

        private void ResetForm()
        {
            WebResourceGrid.ItemsSource = null;
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
            List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            foreach (WebResourceItem webResourceItem in webResources)
            {
                if (webResourceItem.AllowPublish)
                    webResourceItem.Publish = publish;
            }
        }

        private void BoundFile_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Grid grid = (Grid)sender;
            TextBlock textBlock = (TextBlock)grid.Children[0];

            Guid webResourceId = new Guid(textBlock.Tag.ToString());
            FileId.Content = webResourceId;

            List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            WebResourceItem webResourceItem = webResources.FirstOrDefault(w => w.WebResourceId == webResourceId);
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
                ((List<WebResourceItem>)WebResourceGrid.ItemsSource)
                .FirstOrDefault(w => w.WebResourceId == webResourceId);

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

                CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow("Downloaded Web Resource: " + webResource.Id, MessageType.Info);

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

                CrmDeveloperExtensions.Core.FileSystem.WriteFileToDisk(path, decodedContent);

                ProjectItem projectItem = ConnPane.SelectedProject.ProjectItems.AddFromFile(path);

                var fullname = projectItem.FileNames[1];
                var projectPath = System.IO.Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                if (projectPath == null) return;

                var boundName = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                List<WebResourceItem> items = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
                foreach (WebResourceItem item in items.Where(w => w.WebResourceId == webResourceId))
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

            CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow($"Retrieved Web Resource {webResourceId} For Compare", MessageType.Info);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);

            string tempFile = CrmDeveloperExtensions.Core.FileSystem.WriteTempFile(webResource.GetAttributeValue<string>("name"),
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

            List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            if (webResources == null) return;

            webResources.RemoveAll(w => w.WebResourceId == webResourceId);
            //webResources = HandleMappings(webResources);
            WebResourceGrid.ItemsSource = webResources;

            FilterWebResources();

            CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow($"Deleted Web Resource {webResourceId}", MessageType.Info);
        }

        private void ProjectFileList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectFileList.SelectedIndex == -1) return;

            WebResourceItem webResourceItem =
                ((List<WebResourceItem>)WebResourceGrid.ItemsSource)
                .FirstOrDefault(w => w.WebResourceId == new Guid(FileId.Content.ToString()));

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
            {
                MessageBox.Show("Error Retrieving Web Resources. See the Output Window for additional details.");
                return;
            }
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            List<CrmSolution> solutions = Class1.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = solutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private ObservableCollection<MenuItem> GetProjectFolders()
        {
            ObservableCollection<MenuItem> projectFolders = CrmDeveloperExtensions.Core.Vs.ProjectWorker.GetProjectFoldersForMenu(ConnPane.SelectedProject.Name);
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

            CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow("Retrieved Web Resources From CRM", MessageType.Info);

            ObservableCollection<MenuItem> projectFolders = GetProjectFolders();
            List<WebResourceItem> webResourceItems = Class1.CreateWebResourceItemView(results, ConnPane.SelectedProject.Name, projectFolders);
            webResourceItems.ForEach(w => w.PropertyChanged += WebResourceItem_PropertyChanged);

            webResourceItems = webResourceItems.OrderBy(w => w.Name).ToList();

            //webResourceItems = HandleMappings(wrItems);
            WebResourceGrid.ItemsSource = webResourceItems;
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
                    List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
                    foreach (WebResourceItem webResourceItem in webResources.Where(w => w.WebResourceId == item.WebResourceId))
                        webResourceItem.BoundFile = item.BoundFile;
                }

                // AddOrUpdateMapping(item);
            }

            if (e.PropertyName == "Publish")
            {
                List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
                if (webResources == null) return;

                foreach (WebResourceItem webResourceItem in webResources.Where(w => w.WebResourceId == item.WebResourceId))
                    webResourceItem.Publish = item.Publish;

                Publish.IsEnabled = webResources.Count(w => w.Publish) > 0;

                SetPublishAll();
            }
        }

        private void SetPublishAll()
        {
            List<WebResourceItem> webResources = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
            if (webResources == null) return;

            //Set Publish All
            CheckBox publishAll = FindVisualChildren<CheckBox>(WebResourceGrid).FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll == null) return;

            publishAll.IsChecked = webResources.Count(w => w.Publish) == webResources.Count(w => w.AllowPublish);
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
                new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238"); //Default Solution


            //Clear publish flags
            if (!string.IsNullOrEmpty(type))
            {
                List<WebResourceItem> items = (List<WebResourceItem>)WebResourceGrid.ItemsSource;
                foreach (WebResourceItem webResourceItem in items)
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

        private void HideMessage(vsStatusAnimation animation)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        CrmDeveloperExtensions.Core.StatusBar.ClearStatusBarValue(_dte, animation);
                        LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }

        private static bool SetAllowCompare(int type)
        {
            int[] noCompare = { 5, 6, 7, 8, 10 };
            if (!noCompare.Contains(type))
                return true;

            return false;
        }
    }
}
