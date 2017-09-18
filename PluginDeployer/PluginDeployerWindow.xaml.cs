using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using NLog;
using PluginDeployer.ViewModels;
using SparkleXrm.Tasks;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Task = System.Threading.Tasks.Task;

namespace PluginDeployer
{
    public partial class PluginDeployerWindow : INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly EnvDTE.Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private bool _isIlMergeInstalled;
        private ObservableCollection<CrmSolution> _crmSolutions;
        private ObservableCollection<CrmAssembly> _crmAssemblies;

        public ObservableCollection<CrmSolution> CrmSolutions
        {
            get => _crmSolutions;
            set
            {
                _crmSolutions = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<CrmAssembly> CrmAssemblies
        {
            get => _crmAssemblies;
            set
            {
                _crmAssemblies = value;
                OnPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
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

            //WindowEventsOnWindowActivated in this project can be called when activating another window
            //so we don't want to contine further unless our window is active
            if (!HostWindow.IsCrmDevExWindow(gotFocus))
                return;

            //Data is populated already
            //if (SolutionList.ItemsSource != null)
            if (_crmSolutions != null)
                return;

            if (ConnPane.CrmService != null && ConnPane.CrmService.IsReady)
            {
                SetWindowCaption(gotFocus.Caption);
                LoadData();
            }
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            _crmSolutions = new ObservableCollection<CrmSolution>();
            _crmAssemblies = new ObservableCollection<CrmAssembly>();
            LoadData();
            ProjectName.Content = ConnPane.SelectedProject.Name;

            //TODO: better place for this?
            if (!ConfigFile.ConfigFileExists(_dte.Solution.FullName))
                ConfigFile.CreateConfigFile(ConnPane.OrganizationId, ConnPane.SelectedProject.UniqueName, _dte.Solution.FullName);
        }

        private async Task LoadData()
        {
            if (DeploymentType.ItemsSource == null)
            {
                DeploymentType.ItemsSource = AssemblyDeploymentTypes.Types;
                DeploymentType.SelectedIndex = 0;
            }

            await GetCrmData();
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.SetCaption(currentCaption, ConnPane.CrmService);
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            RemoveEventHandlers();

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
            Project project = e.Project;
            if (ConnPane.SelectedProject == project)
                ProjectName.Content = ConnPane.SelectedProject.Name;

            //string solutionPath = Path.GetDirectoryName(_dte.Solution.FullName);
            //if (string.IsNullOrEmpty(solutionPath))
            //    return;

            //string oldName = e.OldName.Replace(solutionPath, string.Empty).Substring(1);

            //Mapping.UpdateProjectName(_dte.Solution.FullName, oldName, project.UniqueName);
        }

        private void ResetForm()
        {
            _crmSolutions = new ObservableCollection<CrmSolution>();
            SolutionList.ItemsSource = null;
            _crmAssemblies = new ObservableCollection<CrmAssembly>();
            CrmAssemblyList.ItemsSource = null;
            DeploymentType.ItemsSource = null;
            ProjectName.Content = string.Empty;
        }

        private async Task GetCrmData()
        {
            Overlay.ShowMessage(_dte, "Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

            RemoveEventHandlers();

            var solutionTask = GetSolutions();
            var assembliesTask = GetCrmAssemblies();

            await Task.WhenAll(solutionTask, assembliesTask);

            SetFormValues();

            FilterAssemblies();

            AddEventHandlers();

            Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);

            if (!solutionTask.Result)
            {
                MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                return;
            }

            if (!assembliesTask.Result)
                MessageBox.Show("Error Retrieving Assemblies. See the Output Window for additional details.");
        }

        private void SetFormValues()
        {
            CrmDevExAssembly crmDevExAssembly = Config.Mapping.HandleMappings(_dte.Solution.FullName,
                ConnPane.SelectedProject, ConnPane.OrganizationId, _crmSolutions, _crmAssemblies);

            if (crmDevExAssembly == null)
            {
                SolutionList.SelectedIndex = 0; //Default
                CrmAssemblyList.SelectedIndex = -1;
                DeploymentType.SelectedIndex = 0;

                return;
            }

            SolutionList.SelectedItem = _crmSolutions.First(s => s.SolutionId == crmDevExAssembly.SolutionId);
            CrmAssemblyList.SelectedItem = _crmAssemblies.First(a => a.AssemblyId == crmDevExAssembly.AssemblyId && a.SolutionId == crmDevExAssembly.SolutionId);

            DeploymentType.SelectedValue = crmDevExAssembly.DeploymentType;
        }

        private async Task<bool> GetSolutions()
        {
            EntityCollection results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Solutions From CRM", MessageType.Info);

            _crmSolutions = ModelBuilder.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = _crmSolutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private async Task<bool> GetCrmAssemblies()
        {
            EntityCollection results = await Task.Run(() => Crm.Assembly.RetrieveAssembliesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            OutputLogger.WriteToOutputWindow("Retrieved Assemblies From CRM", MessageType.Info);

            _crmAssemblies = ModelBuilder.CreateCrmAssemblyView(results);

            CrmAssemblyList.ItemsSource = _crmAssemblies;
            CrmAssemblyList.SelectedIndex = 0;

            return true;
        }

        private async void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            int deploymentType = (int)DeploymentType.SelectedValue;

            switch (deploymentType)
            {
                case 1:
                    await PublishAssemblySpklAsync();
                    break;
                default:
                    await PublishAssemblyAsync();
                    break;
            }
        }

        private async Task PublishAssemblyAsync()
        {
            CrmAssembly crmAssembly;
            if (CrmAssemblyList.SelectedValue != null)
                crmAssembly = (CrmAssembly)CrmAssemblyList.SelectedValue;
            else
                return;

            if (crmAssembly.AssemblyId == Guid.Empty)
                return;

            string path = ProjectWorker.GetOutputFile(ConnPane.SelectedProject);
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Error locating assembly path");
                return;
            }

            bool buildOk = ProjectWorker.BuildProject(ConnPane.SelectedProject);
            if (!buildOk)
                return;

            Version projectAssemblyVersion = Version.Parse(ConnPane.SelectedProject.Properties.Item("AssemblyVersion").Value.ToString());
            bool versionMatch = Versioning.DoAssemblyVersionsMatch(projectAssemblyVersion, crmAssembly.Version);
            if (!versionMatch)
            {
                MessageBox.Show("Error Updating Assembly In CRM: Changes To Major & Minor Versions Require Redeployment");
                return;
            }

            string assemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
            if (!assemblyName.Equals(crmAssembly.Name, StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show("Error Updating Assembly In CRM: Changes To Assembly Name Require Redeployment");
                return;
            }

            bool result;

            try
            {
                Overlay.ShowMessage(_dte, "Deploying...", vsStatusAnimation.vsStatusAnimationDeploy);

                result = await Task.Run(() => Crm.Assembly.UpdateCrmAssembly(ConnPane.CrmService, crmAssembly, path));
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }

            if (!result)
            {
                MessageBox.Show("Error Updating Assembly In CRM: See Output Window for details.");
                return;
            }

            //crmAssembly.Version = projectAssemblyVersion;
            //crmAssembly.Name = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
            //crmAssembly.DisplayName = crmAssembly.Name + " (" + projectAssemblyVersion + ")";
            //crmAssembly.DisplayName += (crmAssembly.IsWorkflow) ? " [Workflow]" : " [Plug-in]";
        }

        private async Task PublishAssemblySpklAsync()
        {
            bool buildOk = ProjectWorker.BuildProject(ConnPane.SelectedProject);
            if (!buildOk)
                return;

            bool isWorkflow = ProjectWorker.IsWorkflowProject(ConnPane.SelectedProject);
            string assemblyPath = ProjectWorker.GetAssemblyPath(ConnPane.SelectedProject);

            bool hasRegistrion = SpklHelpers.RegistrationDetailsPresent(assemblyPath, isWorkflow);
            if (!hasRegistrion)
            {
                MessageBox.Show("You haven't addedd any registration details to the assembly class.");
                return;
            }

            try
            {
                Overlay.ShowMessage(_dte, "Deploying...", vsStatusAnimation.vsStatusAnimationDeploy);

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy;
                var ctx = new OrganizationServiceContext(service);

                CrmSolution solution = (CrmSolution) SolutionList.SelectedItem;
                string solutionName = (solution.SolutionId != ExtensionConstants.DefaultSolutionId)
                    ? solution.UniqueName
                    : null;

                using (ctx)
                {
                    PluginRegistraton pluginRegistraton = new PluginRegistraton(service, ctx, new TraceLogger());

                    Guid assemblyId;
                    if (isWorkflow)
                        assemblyId = await Task.Run(() => pluginRegistraton.RegisterWorkflowActivities(assemblyPath, solutionName));
                    else
                        assemblyId = await Task.Run(() => pluginRegistraton.RegisterPlugin(assemblyPath, solutionName));

                    await SetAssemblyAfterSpklPublish(assemblyId);
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private async Task SetAssemblyAfterSpklPublish(Guid assemblyId)
        {
            Guid crmSolutionId = ((CrmSolution)SolutionList.SelectedValue)?.SolutionId ??
                                 ExtensionConstants.DefaultSolutionId;

            await LoadData();

            CrmAssemblyList.SelectedItem = _crmAssemblies.First(a => a.AssemblyId == assemblyId && a.SolutionId == crmSolutionId);
            DeploymentType.SelectedValue = 1;
        }

        private void SolutionList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateMapping();

            FilterAssemblies();
        }

        private void CrmAssemblyList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateMapping();
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

        private void SpklInstrument_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Instrumenting...", vsStatusAnimation.vsStatusAnimationSync);

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
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void RegistrationTool_OnClick(object sender, RoutedEventArgs e)
        {
            string path = UserOptionsGrid.GetPluginRegistraionToolPath(_dte);

            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Set Plug-in Registraion Tool path under Tools -> Options -> CRM Developer Extensions");
                return;
            }

            if (!path.EndsWith("exe", StringComparison.CurrentCultureIgnoreCase))
                path = Path.Combine(path, "PluginRegistration.exe");


            if (!File.Exists(path))
            {
                MessageBox.Show("PluginRegistration.exe not found at: " + path);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error launching Plug-in Registration Tool: " + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void ConnPane_SelectedProjectChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ConnPane.SelectedProject == null)
                return;

            ProjectName.Content = ConnPane.SelectedProject.Name;
        }

        private void DeploymentType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateMapping();
        }

        private void UpdateMapping()
        {
            CrmAssembly crmAssembly = (CrmAssembly)CrmAssemblyList.SelectedValue;
            if (crmAssembly == null)
                return;

            Guid crmSolutionId = ((CrmSolution)SolutionList.SelectedValue)?.SolutionId ??
                ExtensionConstants.DefaultSolutionId;

            int deploymentType = (int?)DeploymentType.SelectedValue ?? 0;

            Config.Mapping.AddOrUpdateMapping(_dte.Solution.FullName, ConnPane.SelectedProject, crmAssembly.AssemblyId,
                crmSolutionId, deploymentType, ConnPane.OrganizationId);
        }

        private void AddEventHandlers()
        {
            SolutionList.SelectionChanged += SolutionList_OnSelectionChanged;
            CrmAssemblyList.SelectionChanged += CrmAssemblyList_OnSelectionChanged;
            DeploymentType.SelectionChanged += DeploymentType_SelectionChanged;
        }

        private void RemoveEventHandlers()
        {
            SolutionList.SelectionChanged -= SolutionList_OnSelectionChanged;
            CrmAssemblyList.SelectionChanged -= CrmAssemblyList_OnSelectionChanged;
            DeploymentType.SelectionChanged -= DeploymentType_SelectionChanged;
        }

        private void FilterAssemblies()
        {
            Guid crmSolutionId = ((CrmSolution)SolutionList.SelectedValue)?.SolutionId ??
                                 ExtensionConstants.DefaultSolutionId;

            CrmAssembly selectedAssembly = null;
            if (CrmAssemblyList.SelectedItem != null)
                selectedAssembly = (CrmAssembly)CrmAssemblyList.SelectedItem;

            //Apply filter
            ICollectionView icv = CollectionViewSource.GetDefaultView(CrmAssemblyList.ItemsSource);
            if (icv == null) return;
            icv.Filter = o => o is CrmAssembly a && a.SolutionId == crmSolutionId;

            if (selectedAssembly != null)
                CrmAssemblyList.SelectedItem = _crmAssemblies.First(a => a.AssemblyId == selectedAssembly.AssemblyId);
        }
    }
}