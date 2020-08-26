using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.ExtensionMethods;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using NLog;
using PluginDeployer.Resources;
using PluginDeployer.Spkl;
using PluginDeployer.Spkl.Tasks;
using PluginDeployer.ViewModels;
using System;
using System.Collections.Generic;
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
        #region Private

        private readonly DTE _dte;
        private readonly EnvDTE.Solution _solution;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private bool _isIlMergeInstalled;
        private ObservableCollection<CrmSolution> _crmSolutions;
        private ObservableCollection<CrmAssembly> _crmAssemblies;

        #endregion

        #region Public

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

        #endregion

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

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
            //so we don't want to continue further unless our window is active
            if (!HostWindow.IsD365DevExWindow(gotFocus))
                return;

            //Data is populated already
            if (_crmSolutions != null)
                return;

            if (ConnPane.CrmService?.IsReady == true)
                InitializeForm();
        }

        private void InitializeForm()
        {
            ResetCollections();
            LoadData();
            ProjectName.Content = ConnPane.SelectedProject.Name;
            SetWindowCaption(_dte.ActiveWindow.Caption);
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            InitializeForm();
        }

        private void ResetCollections()
        {
            _crmSolutions = new ObservableCollection<CrmSolution>();
            _crmAssemblies = new ObservableCollection<CrmAssembly>();
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
            _dte.ActiveWindow.Caption = HostWindow.GetCaption(currentCaption, ConnPane.CrmService);
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
            var project = e.Project;
            if (ConnPane.SelectedProject == project)
                ResetForm();
        }

        private void ConnPane_OnSolutionProjectRenamed(object sender, SolutionProjectRenamedEventArgs e)
        {
            var project = e.Project;
            if (ConnPane.SelectedProject == project)
                ProjectName.Content = ConnPane.SelectedProject.Name;
        }

        private void ResetForm()
        {
            ResetCollections();
            SolutionList.ItemsSource = null;
            DeploymentType.ItemsSource = null;
            ProjectName.Content = string.Empty;
            BackupFiles.IsChecked = false;
        }

        private async Task GetCrmData()
        {
            ConnPane.CollapsePane();

            var result = false;

            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_RetrievingSolutions}...", vsStatusAnimation.vsStatusAnimationSync);

                var solutionTask = GetSolutions();
                await Task.WhenAll(solutionTask);
                result = solutionTask.Result;
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);

                if (!result)
                    MessageBox.Show(Resource.MessageBox_ErrorRetrievingSolutions);
            }
        }

        private async Task<bool> GetSolutions()
        {
            var results = await Task.Run(() => Crm.Solution.RetrieveSolutionsFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            _crmSolutions = ModelBuilder.CreateCrmSolutionView(results);

            SolutionList.ItemsSource = _crmSolutions;
            SolutionList.SelectedIndex = 0;

            return true;
        }

        private async void Publish_OnClick(object sender, RoutedEventArgs e)
        {
            var deploymentType = (int)DeploymentType.SelectedValue;
            var solution = (CrmSolution)SolutionList.SelectedItem;

            switch (deploymentType)
            {
                case 1:
                    var backupFiles = BackupFiles.ReturnValue();
                    await Task.Run(() => PublishAssemblySpklAsync(solution, backupFiles));
                    break;
                default:
                    await Task.Run(() => PublishAssemblyAsync(solution));
                    break;
            }
        }

        private async Task PublishAssemblyAsync(CrmSolution solution)
        {
            var pluginDeployConfig = Config.Mapping.GetSpklPluginConfig(ConnPane.SelectedProject, ConnPane.SelectedProfile);
            if (!AssemblyValidation.ValidatePluginDeployConfig(pluginDeployConfig))
                return;

            var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
            var assemblyFolderPath = Path.Combine(projectPath, pluginDeployConfig.assemblypath);
            if (!AssemblyValidation.ValidateAssemblyPath(assemblyFolderPath))
                return;

            var assemblyFilePath = Path.Combine(assemblyFolderPath, ConnPane.SelectedProject.Properties.Item("OutputFileName").Value.ToString());

            if (!BuildProject(assemblyFilePath))
                return;

            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_Deploying}...", vsStatusAnimation.vsStatusAnimationDeploy);

                var projectAssemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();

                var isWorkflow = ProjectWorker.IsWorkflowProject(ConnPane.SelectedProject);
                var assemblyProperties = SpklHelpers.AssemblyProperties(assemblyFilePath, isWorkflow);

                var assembly = ModelBuilder.CreateCrmAssembly(projectAssemblyName, assemblyFilePath, assemblyProperties);

                var foundAssembly = Assembly.RetrieveAssemblyFromCrm(ConnPane.CrmService, projectAssemblyName);
                if (foundAssembly != null)
                {
                    var projectAssemblyVersion = Versioning.StringToVersion(assemblyProperties[2]);

                    if (!AssemblyValidation.ValidateAssemblyVersion(ConnPane.CrmService, foundAssembly, projectAssemblyName, projectAssemblyVersion))
                        return;

                    assembly.AssemblyId = foundAssembly.Id;
                }

                var assemblyId = await Task.Run(() => Assembly.UpdateCrmAssembly(ConnPane.CrmService, assembly));
                if (assemblyId == Guid.Empty)
                    MessageBox.Show(Resource.MEssageBox_ErrorDeployingAssembly);

                if (foundAssembly == null)
                    CreatePluginType(assemblyProperties, assemblyId, assemblyFilePath, isWorkflow);

                if (solution.SolutionId == ExtensionConstants.DefaultSolutionId)
                    return;

                var alreadyInSolution = Assembly.IsAssemblyInSolution(ConnPane.CrmService, projectAssemblyName, solution.UniqueName);
                if (alreadyInSolution)
                    return;

                var result = Assembly.AddAssemblyToSolution(ConnPane.CrmService, assemblyId, solution.UniqueName);
                if (!result)
                    MessageBox.Show(Resource.MessageBox_ErrorAddingAssemblyToSolution);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private bool BuildProject(string assemblyFilePath)
        {
            // ReSharper disable once JoinDeclarationAndInitializer
            bool buildResult = false;
#if DEBUG
            ProjectWorker.MovePdbFile(ConnPane.SelectedProject, assemblyFilePath);
            //Initial build fails when debugging because of locked PDB file (which was already moved)
            //Adding a re-build step
            buildResult = ProjectWorker.BuildProject(ConnPane.SelectedProject);
#endif
            if (!buildResult)
                buildResult = ProjectWorker.BuildProject(ConnPane.SelectedProject);

            if (!buildResult)
                OutputLogger.WriteToOutputWindow(Resource.Error_ErrorBuildingProject, MessageType.Error);

            return buildResult;
        }

        private void CreatePluginType(string[] assemblyProperties, Guid assemblyId, string assemblyFilePath, bool isWorkflow)
        {
            var crmPluginRegistrationAttributes = new List<CrmPluginRegistrationAttribute>();
            var crmPluginRegistrationAttribute =
                new CrmPluginRegistrationAttribute(ConnPane.SelectedProject.Name, Guid.NewGuid().ToString(),
                    string.Empty, $"{ConnPane.SelectedProject.Name} ({assemblyProperties[2]})", IsolationModeEnum.Sandbox);

            crmPluginRegistrationAttributes.Add(crmPluginRegistrationAttribute);
            var pluginAssembly = new PluginAssembly { Id = assemblyId };
            var assemblyFullName = SpklHelpers.AssemblyFullName(assemblyFilePath, isWorkflow);

            var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy ?? ConnPane.CrmService.OrganizationWebProxyClient;
            var ctx = new OrganizationServiceContext(service);

            using (ctx)
            {
                var pluginRegistration = new PluginRegistraton(service, ctx, new TraceLogger());

                pluginRegistration.RegisterActivities(crmPluginRegistrationAttributes, pluginAssembly, assemblyFullName);
            }
        }

        private async Task PublishAssemblySpklAsync(CrmSolution solution, bool backupFiles)
        {
            var pluginDeployConfig = Config.Mapping.GetSpklPluginConfig(ConnPane.SelectedProject, ConnPane.SelectedProfile);
            if (!AssemblyValidation.ValidatePluginDeployConfig(pluginDeployConfig))
                return;

            var projectPath = ProjectWorker.GetProjectPath(ConnPane.SelectedProject);
            var assemblyFolderPath = Path.Combine(projectPath, pluginDeployConfig.assemblypath);
            if (!AssemblyValidation.ValidateAssemblyPath(assemblyFolderPath))
                return;

            var assemblyFilePath = Path.Combine(assemblyFolderPath, ConnPane.SelectedProject.Properties.Item("OutputFileName").Value.ToString());

            if (!BuildProject(assemblyFilePath))
                return;

            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_Deploying}...", vsStatusAnimation.vsStatusAnimationDeploy);

                var isWorkflow = ProjectWorker.IsWorkflowProject(ConnPane.SelectedProject);
                if (!AssemblyValidation.ValidateRegistrationDetails(assemblyFilePath, isWorkflow))
                    return;

                var assemblyProperties = SpklHelpers.AssemblyProperties(assemblyFilePath, isWorkflow);
                var projectAssemblyVersion = Version.Parse(assemblyProperties[2]);

                var projectAssemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
                var foundAssembly = Assembly.RetrieveAssemblyFromCrm(ConnPane.CrmService, projectAssemblyName);
                if (foundAssembly != null)
                {
                    if (!AssemblyValidation.ValidateAssemblyVersion(ConnPane.CrmService, foundAssembly, projectAssemblyName, projectAssemblyVersion))
                        return;
                }

                var solutionName = solution.SolutionId != ExtensionConstants.DefaultSolutionId
                    ? solution.UniqueName
                    : null;

                var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy ?? ConnPane.CrmService.OrganizationWebProxyClient;
                var ctx = new OrganizationServiceContext(service);

                using (ctx)
                {
                    var pluginRegistration = new PluginRegistraton(service, ctx, new TraceLogger());

                    if (isWorkflow)
                        await Task.Run(() => pluginRegistration.RegisterWorkflowActivities(assemblyFilePath, solutionName));
                    else
                        await Task.Run(() => pluginRegistration.RegisterPlugin(assemblyFilePath, solutionName));

                    GetRegistrationDetailsWithContext(pluginDeployConfig.classRegex, backupFiles, ctx);
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void IlMerge_OnClick(object sender, RoutedEventArgs e)
        {
            if (ConnPane.SelectedProject == null)
                return;

            _isIlMergeInstalled = IlMergeHandler.ToggleIlMerge(ConnPane.SelectedProject, _isIlMergeInstalled);
        }

        private void AddRegistration_OnClick(object sender, RoutedEventArgs e)
        {
            var pluginDeployConfig = Config.Mapping.GetSpklPluginConfig(ConnPane.SelectedProject, ConnPane.SelectedProfile);
            if (pluginDeployConfig == null)
            {
                MessageBox.Show($"{Resource.MessageBox_MissingPluginsSpklConfig}: {ExtensionConstants.SpklConfigFile}");
                return;
            }

            var hasRegAttributeClass = SpklHelpers.RegAttributeDefinitionExists(ConnPane.SelectedProject);
            if (!hasRegAttributeClass)
            {
                //TODO: If VB support is added this would need to be addressed
                TemplateHandler.AddFileFromTemplate(ConnPane.SelectedProject,
                    "CSharpSpklRegAttributes\\CSharpSpklRegAttributes", $"{ExtensionConstants.SpklRegAttrClassName}.cs");
            }

            GetRegistrationDetailsWithoutContext(pluginDeployConfig.classRegex, BackupFiles.ReturnValue());
        }

        private void GetRegistrationDetailsWithContext(string customClassRegex, bool backupFiles, OrganizationServiceContext ctx)
        {
            try
            {
                Overlay.ShowMessage(_dte, $"{Resource.Message_AddingRegistrationDetails}...", vsStatusAnimation.vsStatusAnimationSync);

                var project = ConnPane.SelectedProject;
                ProjectWorker.BuildProject(project);

                var path = Path.GetDirectoryName(project.FullName);

                var downloadPluginMetadataTask = new DownloadPluginMetadataTask(ctx, new TraceLogger());
                downloadPluginMetadataTask.Execute(path, backupFiles, customClassRegex);
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void GetRegistrationDetailsWithoutContext(string customClassRegex, bool backupFiles)
        {
            var service = (IOrganizationService)ConnPane.CrmService.OrganizationServiceProxy ?? ConnPane.CrmService.OrganizationWebProxyClient;
            var ctx = new OrganizationServiceContext(service);

            GetRegistrationDetailsWithContext(customClassRegex, backupFiles, ctx);
        }

        private void RegistrationTool_OnClick(object sender, RoutedEventArgs e)
        {
            PrtHelper.OpenPrt();
        }

        private void ConnPane_SelectedProjectChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ConnPane.SelectedProject == null)
                return;

            ProjectName.Content = ConnPane.SelectedProject.Name;
        }

        private void OpenInCrm_Click(object sender, RoutedEventArgs e)
        {
            CrmSolution solution = (CrmSolution)SolutionList.SelectedItem;

            D365DeveloperExtensions.Core.WebBrowser.OpenCrmPage(ConnPane.CrmService, $"tools/solution/edit.aspx?id=%7b{solution.SolutionId}%7d");
        }
    }
}