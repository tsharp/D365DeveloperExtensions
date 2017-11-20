using CrmDeveloperExtensions2.Core;
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
using PluginDeployer.Spkl;
using PluginDeployer.Spkl.Tasks;
using PluginDeployer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Assembly = PluginDeployer.Crm.Assembly;
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

            SetWindowCaption(_dte.ActiveWindow.Caption);
        }

        private async void LoadData()
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
        }

        private void ResetForm()
        {
            _crmSolutions = new ObservableCollection<CrmSolution>();
            SolutionList.ItemsSource = null;
            _crmAssemblies = new ObservableCollection<CrmAssembly>();
            DeploymentType.ItemsSource = null;
            ProjectName.Content = string.Empty;
            BackupFiles.IsChecked = false;
        }

        private async Task GetCrmData()
        {
            bool result = false;

            try
            {
                Overlay.ShowMessage(_dte, "Getting CRM data...", vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();
                await Task.WhenAll(solutionTask);
                result = solutionTask.Result;
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);

                if (!result)
                    MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
            }
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

        private async void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            int deploymentType = (int)DeploymentType.SelectedValue;

            switch (deploymentType)
            {
                case 1:
                    await PublishAssemblySpklAsync();
                    break;
                default:
                    CrmSolution solution = (CrmSolution)SolutionList.SelectedItem;
                    await Task.Run(() => PublishAssemblyAsync(solution));
                    break;
            }
        }

        private async Task PublishAssemblyAsync(CrmSolution solution)
        {
            bool buildOk = ProjectWorker.BuildProject(ConnPane.SelectedProject);
            if (!buildOk)
                return;

            try
            {
                Overlay.ShowMessage(_dte, "Deploying...", vsStatusAnimation.vsStatusAnimationDeploy);

                string projectAssemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
                string assemblyFolderPath = ProjectWorker.GetOutputPath(ConnPane.SelectedProject);
                if (!SpklHelpers.ValidateAssemblyPath(assemblyFolderPath))
                    return;

                string assemblyFilePath = ProjectWorker.GetAssemblyPath(ConnPane.SelectedProject);
                string[] assemblyProperties = SpklHelpers.AssemblyProperties(assemblyFilePath);

                CrmAssembly assembly = new CrmAssembly
                {
                    Name = projectAssemblyName,
                    AssemblyPath = assemblyFilePath,
                    Version = assemblyProperties[2],
                    Culture = assemblyProperties[4],
                    PublicKeyToken = assemblyProperties[6],
                    //TODO: option to make none?
                    IsolationMode = IsolationModeEnum.Sandbox
                };

                Entity foundAssembly = Assembly.RetrieveAssemblyFromCrm(ConnPane.CrmService, projectAssemblyName);
                if (foundAssembly != null)
                {

                    Version projectAssemblyVersion = Versioning.StringToVersion(assemblyProperties[2]);

                    if (!SpklHelpers.ValidateAssemblyVersion(ConnPane.CrmService, foundAssembly, projectAssemblyName, projectAssemblyVersion))
                        return;

                    assembly.AssemblyId = foundAssembly.Id;
                }

                Guid assemblyId = await Task.Run(() => Assembly.UpdateCrmAssembly(ConnPane.CrmService, assembly));
                if (assemblyId == Guid.Empty)
                    MessageBox.Show("Error Deploying Assembly In CRM: See Output Window for details.");

                if (solution.SolutionId == ExtensionConstants.DefaultSolutionId)
                    return;

                bool alreadyInSolution = Assembly.IsAssemblyInSolution(ConnPane.CrmService, projectAssemblyName, solution.UniqueName);
                if (alreadyInSolution)
                    return;

                bool result = Assembly.AddAssemblyToSolution(ConnPane.CrmService, assemblyId, solution.UniqueName);
                if (!result)
                    MessageBox.Show("Error Adding Assembly To Solution In CRM: See Output Window for details.");
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private async Task PublishAssemblySpklAsync()
        {
            bool buildOk = ProjectWorker.BuildProject(ConnPane.SelectedProject);
            if (!buildOk)
                return;

            try
            {
                Overlay.ShowMessage(_dte, "Deploying...", vsStatusAnimation.vsStatusAnimationDeploy);

                string projectAssemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();

                PluginDeployConfig pluginDeployConfig = Config.Mapping.GetSpklPluginConfig(ConnPane.SelectedProject, ConnPane.SelectedProfile);
                if (pluginDeployConfig == null)
                {
                    MessageBox.Show($"Missing 'plugins' configuration in {ExtensionConstants.SpklConfigFile}");
                    return;
                }

                if (string.IsNullOrEmpty(pluginDeployConfig.assemblypath))
                {
                    MessageBox.Show($"Missing 'assemblypath' in 'plugins' configuration in {ExtensionConstants.SpklConfigFile}");
                    return;
                }

                string projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
                string assemblyFolderPath = Path.Combine(projectPath, pluginDeployConfig.assemblypath);
                if (!SpklHelpers.ValidateAssemblyPath(assemblyFolderPath))
                    return;

                bool isWorkflow = ProjectWorker.IsWorkflowProject(ConnPane.SelectedProject);
                string assemblyFilePath = Path.Combine(assemblyFolderPath, ConnPane.SelectedProject.Properties.Item("OutputFileName").Value.ToString());
                if (!SpklHelpers.ValidateRegistraionDetails(assemblyFilePath, isWorkflow))
                    return;

                string[] assemblyProperties = SpklHelpers.AssemblyProperties(assemblyFilePath);
                Version projectAssemblyVersion = Version.Parse(assemblyProperties[2]);

                Entity foundAssembly = Assembly.RetrieveAssemblyFromCrm(ConnPane.CrmService, projectAssemblyName);
                if (foundAssembly != null)
                {
                    if (!SpklHelpers.ValidateAssemblyVersion(ConnPane.CrmService, foundAssembly, projectAssemblyName, projectAssemblyVersion))
                        return;
                }

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy;
                var ctx = new OrganizationServiceContext(service);

                CrmSolution solution = (CrmSolution)SolutionList.SelectedItem;
                string solutionName = solution.SolutionId != ExtensionConstants.DefaultSolutionId
                    ? solution.UniqueName
                    : null;

                using (ctx)
                {
                    PluginRegistraton pluginRegistraton = new PluginRegistraton(service, ctx, new TraceLogger());

                    if (isWorkflow)
                        await Task.Run(() => pluginRegistraton.RegisterWorkflowActivities(assemblyFilePath, solutionName));
                    else
                        await Task.Run(() => pluginRegistraton.RegisterPlugin(assemblyFilePath, solutionName));

                    GetRegistrationDetails(pluginDeployConfig.classRegex);
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
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

        private void AddRegistration_OnClick(object sender, RoutedEventArgs e)
        {
            PluginDeployConfig pluginDeployConfig = Config.Mapping.GetSpklPluginConfig(ConnPane.SelectedProject, ConnPane.SelectedProfile);
            if (pluginDeployConfig == null)
            {
                MessageBox.Show($"Missing 'plugins' configuration in {ExtensionConstants.SpklConfigFile}");
                return;
            }

            GetRegistrationDetails(pluginDeployConfig.classRegex);
        }

        private void GetRegistrationDetails(string customClassRegex)
        {
            try
            {
                Overlay.ShowMessage(_dte, "Adding Registration Details...", vsStatusAnimation.vsStatusAnimationSync);

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy;
                var ctx = new OrganizationServiceContext(service);

                Project project = ConnPane.SelectedProject;
                ProjectWorker.BuildProject(project);

                string path = Path.GetDirectoryName(project.FullName);
                bool backupFiles = BackupFiles.IsChecked.HasValue && BackupFiles.IsChecked.Value;

                DownloadPluginMetadataTask downloadPluginMetadataTask = new DownloadPluginMetadataTask(ctx, new TraceLogger());
                downloadPluginMetadataTask.Execute(path, backupFiles, customClassRegex);
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

        private void SpklHelp_Click(object sender, RoutedEventArgs e)
        {
            CrmDeveloperExtensions2.Core.WebBrowser.OpenUrl(_dte, "https://github.com/scottdurow/SparkleXrm/wiki/spkl");
        }

        private void OpenInCrm_Click(object sender, RoutedEventArgs e)
        {
            CrmSolution solution = (CrmSolution)SolutionList.SelectedItem;

            CrmDeveloperExtensions2.Core.WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService, $"tools/solution/edit.aspx?id=%7b{solution.SolutionId}%7d");
        }
    }
}