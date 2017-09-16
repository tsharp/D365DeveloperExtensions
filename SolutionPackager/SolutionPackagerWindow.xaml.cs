using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using NLog;
using SolutionPackager.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;
using Window = EnvDTE.Window;

namespace SolutionPackager
{
    public partial class SolutionPackagerWindow : INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private ObservableCollection<CrmSolution> _solutionData;
        private ObservableCollection<string> _solutionFolders;

        public bool SolutionXmlExists;
        public ObservableCollection<CrmSolution> SolutionData
        {
            get => _solutionData;
            set
            {
                _solutionData = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<string> SolutionFolders
        {
            get => _solutionFolders;
            set
            {
                _solutionFolders = value;
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

            SolutionFolders = new ObservableCollection<string>();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;

            DataObject.AddPastingHandler(VersionMajor, TextBoxPasting);
            DataObject.AddPastingHandler(VersionMinor, TextBoxPasting);
            DataObject.AddPastingHandler(VersionBuild, TextBoxPasting);
            DataObject.AddPastingHandler(VersionRevision, TextBoxPasting);
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
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

            //Window was already loaded
            if (SolutionData != null)
                return;

            if (ConnPane.CrmService != null && ConnPane.CrmService.IsReady)
            {
                SetWindowCaption(gotFocus.Caption);
                SetControlState(true);
                BindPackageButton();
                LoadData();
            }
        }

        private async void LoadData()
        {
            GetSolutionFolders();
            await GetCrmData();
        }

        private void BindPackageButton()
        {
            SolutionXmlExists = SolutionXml.SolutionXmlExists(ConnPane.SelectedProject) && SolutionData != null;
        }

        private void GetSolutionFolders()
        {
            SolutionFolders = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjectFolders(ConnPane.SelectedProject);
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.SetCaption(currentCaption, ConnPane.CrmService);
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            SetControlState(true);

            LoadData();

            if (!ConfigFile.ConfigFileExists(_dte.Solution.FullName))
                ConfigFile.CreateConfigFile(ConnPane.OrganizationId, ConnPane.SelectedProject.UniqueName, _dte.Solution.FullName);

            SetWindowCaption(_dte.ActiveWindow.Caption);
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

        private void ConnPane_OnSolutionProjectRemoved(object sender, SolutionProjectRemovedEventArgs e)
        {
            Project project = e.Project;
            if (ConnPane.SelectedProject == project)
                ResetForm();
        }

        private void ConnPane_OnSolutionProjectRenamed(object sender, SolutionProjectRenamedEventArgs e)
        {
            //Project project = e.Project;
            //string solutionPath = Path.GetDirectoryName(_dte.Solution.FullName);
            //if (string.IsNullOrEmpty(solutionPath))
            //    return;

            //string oldName = e.OldName.Replace(solutionPath, string.Empty).Substring(1);

            //Mapping.UpdateProjectName(_dte.Solution.FullName, oldName, project.UniqueName);
        }

        private void SetControlState(bool enabled)
        {
            if (enabled == false)
                PackageSolution.IsEnabled = false;
            //UnpackageSolution.IsEnabled = enabled;
            SolutionList.IsEnabled = enabled;
            SaveSolutions.IsEnabled = enabled;
            ProjectFolder.IsEnabled = enabled;
            EnableSolutionPackagerLog.IsEnabled = enabled;
            DownloadManaged.IsEnabled = enabled;
            CreateManaged.IsEnabled = enabled;
            VersionMajor.IsEnabled = enabled;
            VersionMinor.IsEnabled = enabled;
            VersionBuild.IsEnabled = enabled;
            VersionRevision.IsEnabled = enabled;
            UpdateVersion.IsEnabled = enabled;
            PublishAll.IsEnabled = enabled;
        }

        private void ResetForm()
        {
            RemoveEventHandlers();
            SolutionData = null;
            SolutionFolders = new ObservableCollection<string>();
            SetControlState(false);
        }

        private async Task GetCrmData()
        {
            try
            {
                Overlay.ShowMessage(_dte, "Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();

                await Task.WhenAll(solutionTask);

                if (!solutionTask.Result)
                {
                    Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                }

                AddEventHandlers();
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
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

            CrmDevExSolutionPackage crmDevExSolutionPackage =
                Config.Mapping.HandleMappings(_dte.Solution.FullName, ConnPane.SelectedProject,
                    SolutionData, ConnPane.OrganizationId);

            if (crmDevExSolutionPackage == null)
            {
                ProjectFolder.SelectedItem = SolutionFolders.FirstOrDefault(s => s == "/_Solutions");
                return true;
            }

            SetControlStateForItem(crmDevExSolutionPackage);

            return true;
        }

        private void SetControlStateForItem(CrmDevExSolutionPackage crmDevExSolutionPackage)
        {
            SolutionList.SelectedItem = SolutionData.FirstOrDefault(s => s.SolutionId == crmDevExSolutionPackage.SolutionId);
            SaveSolutions.IsChecked = crmDevExSolutionPackage.SaveSolutions;
            ProjectFolder.SelectedItem = SolutionFolders.FirstOrDefault(s => s == crmDevExSolutionPackage.ProjectFolder);
            EnableSolutionPackagerLog.IsChecked = crmDevExSolutionPackage.EnableSolutionPackagerLog;
            DownloadManaged.IsChecked = crmDevExSolutionPackage.DownloadManaged;
            CreateManaged.IsChecked = crmDevExSolutionPackage.CreateManaged;
            PublishAll.IsChecked = crmDevExSolutionPackage.PublishAll;

            PackageSolution.IsEnabled = SolutionXml.SolutionXmlExists(ConnPane.SelectedProject);
            if (PackageSolution.IsEnabled)
                SetFormVersionNumbers();
        }

        private void PublishSolution_OnClick(object sender, RoutedEventArgs e)
        {
            PublishSolutionToCrm();
        }

        private async void PublishSolutionToCrm()
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to publish solution?", "Ok to publish?",
                MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);

            if (result == MessageBoxResult.No)
                return;

            string latestSolutionPath =
                SolutionXml.GetLatestSolutionPath(ConnPane.SelectedProject, ProjectFolder.SelectedItem.ToString());

            if (string.IsNullOrEmpty(latestSolutionPath))
            {
                MessageBox.Show("Unable to find latest solution.");
                return;
            }
            bool publishAll = PublishAll.IsChecked == true;
            bool success;
            try
            {
                Overlay.ShowMessage(_dte, "Importing solution...", vsStatusAnimation.vsStatusAnimationDeploy);

                success = await Task.Run(() => PublishToCrm(latestSolutionPath, publishAll));
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
            if (!success)
                MessageBox.Show("Error importing or publishing solution. See output window for details.");

        }

        private async Task<bool> PublishToCrm(string latestSolutionPath, bool publishAll)
        {
            var success = await Task.Run(() => Crm.Solution.ImportSolution(ConnPane.CrmService, latestSolutionPath));

            if (!publishAll)
                return success;

            Overlay.ShowMessage(_dte, "Publishing customizations...", vsStatusAnimation.vsStatusAnimationDeploy);

            success =
                await Task.Run(() => CrmDeveloperExtensions2.Core.Crm.Publish.PublishAllCustomizations(ConnPane.CrmService));

            return success;
        }

        private CrmDevExSolutionPackage CreateMappingObject()
        {
            if (SolutionList.SelectedItem == null)
                return null;

            return new CrmDevExSolutionPackage
            {
                SolutionId = ((CrmSolution)SolutionList.SelectedItem).SolutionId,
                SaveSolutions = SaveSolutions.IsChecked ?? false,
                EnableSolutionPackagerLog = EnableSolutionPackagerLog.IsChecked ?? false,
                ProjectFolder = ProjectFolder.SelectedItem.ToString(),
                DownloadManaged = DownloadManaged.IsChecked ?? false,
                CreateManaged = CreateManaged.IsChecked ?? false,
                PublishAll = PublishAll.IsChecked ?? false
            };
        }

        private void AddEventHandlers()
        {
            SolutionList.SelectionChanged += SolutionList_OnSelectionChanged;
            DownloadManaged.Checked += DownloadManaged_OnChecked;
            DownloadManaged.Unchecked += DownloadManaged_OnChecked;
            SaveSolutions.Checked += SaveSolutions_OnChecked;
            SaveSolutions.Unchecked += SaveSolutions_OnChecked;
            CreateManaged.Checked += CreateManaged_OnChecked;
            CreateManaged.Unchecked += CreateManaged_OnChecked;
            PublishAll.Checked += PublishAll_OnChecked;
            PublishAll.Unchecked += PublishAll_OnChecked;
        }

        private void RemoveEventHandlers()
        {
            SolutionList.SelectionChanged -= SolutionList_OnSelectionChanged;
            DownloadManaged.Checked -= DownloadManaged_OnChecked;
            DownloadManaged.Unchecked -= DownloadManaged_OnChecked;
            SaveSolutions.Checked -= SaveSolutions_OnChecked;
            SaveSolutions.Unchecked -= SaveSolutions_OnChecked;
            CreateManaged.Checked -= CreateManaged_OnChecked;
            CreateManaged.Unchecked -= CreateManaged_OnChecked;
            PublishAll.Checked -= PublishAll_OnChecked;
            PublishAll.Unchecked -= PublishAll_OnChecked;
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!SolutionList.IsLoaded)
                return;

            if (SolutionList.SelectedItem == null)
            {
                SetControlState(false);

                Config.Mapping.AddOrUpdateMapping(_dte.Solution.FullName, ConnPane.SelectedProject, null, ConnPane.OrganizationId);
                return;
            }

            SetControlState(true);
            SetFormVersionNumbers();
            Config.Mapping.AddOrUpdateMapping(_dte.Solution.FullName, ConnPane.SelectedProject, CreateMappingObject(), ConnPane.OrganizationId);
        }

        private void TriggerMappingUpdate(object sender)
        {
            Control c = (Control)sender;
            if (!c.IsLoaded)
                return;

            Config.Mapping.AddOrUpdateMapping(_dte.Solution.FullName, ConnPane.SelectedProject, CreateMappingObject(), ConnPane.OrganizationId);
        }

        private void DownloadManaged_OnChecked(object sender, RoutedEventArgs e)
        {
            TriggerMappingUpdate(sender);
        }

        private void SaveSolutions_OnChecked(object sender, RoutedEventArgs e)
        {
            TriggerMappingUpdate(sender);
        }

        private void CreateManaged_OnChecked(object sender, RoutedEventArgs e)
        {
            TriggerMappingUpdate(sender);
        }

        private void PublishAll_OnChecked(object sender, RoutedEventArgs e)
        {
            TriggerMappingUpdate(sender);
        }

        private void ConnPane_OnSelectedProjectChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox solutionProjectsList = (ComboBox)e.Source;
            if (!solutionProjectsList.IsLoaded || ConnPane.SelectedProject == null)
                return;

            Config.Mapping.HandleMappings(_dte.Solution.FullName, ConnPane.SelectedProject, SolutionData,
                ConnPane.OrganizationId);
        }

        private void PackageSolution_OnClick(object sender, RoutedEventArgs e)
        {
            PackageProcess();
        }

        private void PackageProcess()
        {
            try
            {
                if (SolutionList.SelectedItem == null)
                    return;

                Version version = SolutionXml.GetSolutionXmlVersion(ConnPane.SelectedProject);
                if (version == null)
                {
                    MessageBox.Show("Invalid Solution.xml version number. See the Output Window for additional details.");
                    return;
                }

                CrmSolution selectedSolution = (CrmSolution)SolutionList.SelectedItem;
                CrmDevExSolutionPackage crmDevExSolutionPackage = CreateMappingObject();

                Overlay.ShowMessage(_dte, "Packaging solution...", vsStatusAnimation.vsStatusAnimationSync);

                bool success = Packager.CreatePackage(_dte, selectedSolution.UniqueName, version, ConnPane.SelectedProject, crmDevExSolutionPackage);

                if (!success)
                {
                    Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Packaging Solution. See the Output Window for additional details.");
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void UnpackageSolution_OnClick(object sender, RoutedEventArgs e)
        {
            UnpackageProcess();
        }

        private async void UnpackageProcess()
        {
            try
            {
                if (SolutionList.SelectedItem == null)
                    return;

                CrmSolution selectedSolution = (CrmSolution)SolutionList.SelectedItem;
                CrmDevExSolutionPackage crmDevExSolutionPackage = CreateMappingObject();

                Overlay.ShowMessage(_dte, "Connecting to CRM/365 and getting unmanaged solution...", vsStatusAnimation.vsStatusAnimationSync);

                List<Task> tasks = new List<Task>();
                var getUmanagedSolution = Crm.Solution.GetSolutionFromCrm(ConnPane.CrmService, selectedSolution, false);
                tasks.Add(getUmanagedSolution);

                Task<string> getManagedSolution = null;
                if (crmDevExSolutionPackage.DownloadManaged)
                {
                    getManagedSolution = Crm.Solution.GetSolutionFromCrm(ConnPane.CrmService, selectedSolution, true);
                    tasks.Add(getManagedSolution);
                }

                await Task.WhenAll(tasks);

                if (string.IsNullOrEmpty(getUmanagedSolution.Result))
                {
                    Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Retrieving Unmanaged Solution. See the Output Window for additional details.");
                    return;
                }

                if (crmDevExSolutionPackage.DownloadManaged && getManagedSolution != null)
                {
                    if (string.IsNullOrEmpty(getManagedSolution.Result))
                    {
                        Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                        MessageBox.Show("Error Retrieving Managed Solution. See the Output Window for additional details.");
                        return;
                    }
                }

                OutputLogger.WriteToOutputWindow("Retrieved Unmanaged Solution From CRM", MessageType.Info);
                Overlay.ShowMessage(_dte, "Extracting solution...", vsStatusAnimation.vsStatusAnimationSync);

                bool success = Packager.ExtractPackage(_dte, getUmanagedSolution.Result, getManagedSolution?.Result, ConnPane.SelectedProject, crmDevExSolutionPackage);

                if (!success)
                    MessageBox.Show("Error Extracting Solution. See the Output Window for additional details.");

                PackageSolution.IsEnabled = true;
                SetFormVersionNumbers();
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void ConnPane_OnProjectItemAdded(object sender, ProjectItemAddedEventArgs e)
        {
            BindPackageButton();

            ProjectItem projectItem = e.ProjectItem;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            string newItemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);
            SolutionFolders.Add(newItemName);

            SolutionFolders = new ObservableCollection<string>(SolutionFolders.OrderBy(s => s));
        }

        private void ConnPane_OnProjectItemRemoved(object sender, ProjectItemRemovedEventArgs e)
        {
            BindPackageButton();

            ProjectItem projectItem = e.ProjectItem;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return;

            var itemName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            SolutionFolders.Remove(itemName);

            SolutionFolders = new ObservableCollection<string>(SolutionFolders.OrderBy(s => s));
        }

        private void ConnPane_OnProjectItemRenamed(object sender, ProjectItemRenamedEventArgs e)
        {
            BindPackageButton();

            ProjectItem projectItem = e.ProjectItem;
            if (projectItem.Name == null) return;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;
            string oldName = e.OldName;
            Guid itemType = new Guid(projectItem.Kind);

            if (itemType != VSConstants.GUID_ItemType_PhysicalFolder)
                return;

            var newItemPath = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);

            int index = newItemPath.LastIndexOf(projectItem.Name, StringComparison.Ordinal);
            if (index == -1) return;

            var oldItemPath = newItemPath.Remove(index, projectItem.Name.Length).Insert(index, oldName);

            SolutionFolders.Remove(oldItemPath);

            SolutionFolders = new ObservableCollection<string>(SolutionFolders.OrderBy(s => s));
        }

        private void SetFormVersionNumbers()
        {
            Version version = SolutionXml.GetSolutionXmlVersion(ConnPane.SelectedProject);
            if (version == null)
            {
                VersionMajor.Text = String.Empty;
                VersionMinor.Text = String.Empty;
                VersionBuild.Text = String.Empty;
                VersionRevision.Text = String.Empty;
                return;
            }

            VersionMajor.Text = version.Major.ToString();
            VersionMinor.Text = version.Minor.ToString();
            VersionBuild.Text = version.Build != -1 ? version.Build.ToString() : String.Empty;
            VersionRevision.Text = version.Revision != -1 ? version.Revision.ToString() : String.Empty;
        }

        private void UpdateVersion_OnClick(object sender, RoutedEventArgs e)
        {
            Version version = ValidateVersionInput(VersionMajor.Text, VersionMinor.Text,
                VersionBuild.Text, VersionRevision.Text);
            if (version == null)
            {
                MessageBox.Show("Invalid version number");
                return;
            }

            bool success = SolutionXml.SetSolutionXmlVersion(ConnPane.SelectedProject, version);
            if (!success)
                MessageBox.Show("Error updating Solution.xml version: see output window for details");
        }

        private Version ValidateVersionInput(string majorIn, string minorIn, string buildIn, string revisionIn)
        {
            bool isMajorInt = int.TryParse(majorIn, out int major);
            bool isMinorInt = int.TryParse(minorIn, out int minor);
            bool isBuildInt = int.TryParse(buildIn, out int build);
            bool isRevisionInt = int.TryParse(revisionIn, out int revision);

            if (!isMajorInt || !isMinorInt)
                return null;

            string v = string.Concat(major, ".", minor, (isBuildInt) ? $".{build}" : null,
                isRevisionInt ? $".{revision}" : null);

            return new Version(v);
        }

        private void Version_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            Regex regex = new Regex("[^0-9.-]+");
            return !regex.IsMatch(text);
        }

        private void TextBoxPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsTextAllowed(text))
                    e.CancelCommand();
            }
            else
                e.CancelCommand();
        }
    }
}