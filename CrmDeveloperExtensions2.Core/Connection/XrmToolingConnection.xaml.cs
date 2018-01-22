using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.ExtensionMethods;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Resources;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using Window = EnvDTE.Window;

namespace CrmDeveloperExtensions2.Core.Connection
{
    public partial class XrmToolingConnection : INotifyPropertyChanged
    {
        #region Private

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly DTE _dte;
        private readonly Solution _solution;
        private ObservableCollection<ProjectListItem> _projects;
        private List<string> _profiles;
        private bool _projectEventsRegistered;
        private IVsSolution _vsSolution;
        private IVsSolutionEvents _vsSolutionEvents;
        private readonly List<MovedProjectItem> _movedProjectItems;
        private bool _autoLogin;
        private bool _isConnected;
        private Project _selectedProject;
        private string _selectedProfile;

        #endregion

        #region Public

        public CrmServiceClient CrmService;
        public Guid OrganizationId;
        public ObservableCollection<ProjectListItem> Projects
        {
            get => _projects;
            set
            {
                if (_projects == value)
                    return;

                _projects = value;
                OnPropertyChanged();
            }
        }
        public Project SelectedProject
        {
            get => _selectedProject;
            set
            {
                _selectedProject = value;
                OnPropertyChanged();
            }
        }
        public string SelectedProfile
        {
            get => _selectedProfile;
            set
            {
                _selectedProfile = value;
                OnPropertyChanged();
            }
        }
        public bool IsConnected
        {
            get => _isConnected;
            set
            {
                _isConnected = value;
                OnPropertyChanged();
            }
        }
        public List<string> Profiles
        {
            get => _profiles;
            set
            {
                _profiles = value;
                OnPropertyChanged();
            }
        }
        public ToolWindow ToolWindow;

        #endregion

        #region Event Handlers

        public event EventHandler<ConnectEventArgs> Connected;
        public event EventHandler SolutionOpened;
        public event EventHandler SolutionBeforeClosing;
        public event EventHandler<SolutionProjectAddedEventArgs> SolutionProjectAdded;
        public event EventHandler<SolutionProjectRemovedEventArgs> SolutionProjectRemoved;
        public event EventHandler<SolutionProjectRenamedEventArgs> SolutionProjectRenamed;
        public event EventHandler<ProjectItemRemovedEventArgs> ProjectItemRemoved;
        public event EventHandler<ProjectItemAddedEventArgs> ProjectItemAdded;
        public event EventHandler<ProjectItemRenamedEventArgs> ProjectItemRenamed;
        public event EventHandler<ProjectItemMovedEventArgs> ProjectItemMoved;
        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler<SelectionChangedEventArgs> SelectedProjectChanged;
        public event EventHandler<SelectionChangedEventArgs> ProfileChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        public XrmToolingConnection()
        {
            InitializeComponent();
            DataContext = this;

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            AdviseSolutionEvents();

            var events = _dte.Events;
            BindWindowEvents(events);

            BindProjectItemEvents();

            BindSolutionEvents(events);

            _movedProjectItems = new List<MovedProjectItem>();
            Projects = new ObservableCollection<ProjectListItem>();
            Profiles = new List<string>();
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

            ToolWindow = HostWindow.GetCrmDevExWindow(_dte.ActiveWindow);

            if (Projects == null || Projects.Count == 0)
            {
                GetProjectsForList();
                SolutionProjectsList.SelectionChanged += SolutionProjectsList_OnSelectionChanged;

                SetConfigFile();
            }

            if (Profiles == null || Profiles.Count == 0)
                GetProfiles();

            if (!_projectEventsRegistered)
            {
                RegisterProjectEvents();
                _projectEventsRegistered = true;
            }

            if (CrmService != null)
                return;

            CrmServiceClient client = SharedGlobals.GetGlobal("CrmService", _dte) as CrmServiceClient;
            if (client?.ConnectedOrgUniqueName == null)
                return;

            CrmService = client;
            _isConnected = true;
        }

        #region Event Binding
        private void AdviseSolutionEvents()
        {
            _vsSolutionEvents = new VsSolutionEvents(_dte, this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out uint solutionEventsCookie);
        }

        private void BindSolutionEvents(Events events)
        {
            var solutionEvents = events.SolutionEvents;
            solutionEvents.ProjectAdded += SolutionEventsOnProjectAdded;
            solutionEvents.ProjectRemoved += SolutionEventsOnProjectRemoved;
            solutionEvents.ProjectRenamed += SolutionEventsOnProjectRenamed;
            solutionEvents.BeforeClosing += SolutionEventsOnBeforeClosing;
            solutionEvents.Opened += SolutionEventsOnOpened;
        }

        private void BindProjectItemEvents()
        {
            var events2 = (Events2)_dte.Events;
            var projectItemsEvents = events2.ProjectItemsEvents;
            projectItemsEvents.ItemRenamed += ProjectItemsEventsOnItemRenamed;
            projectItemsEvents.ItemAdded += ProjectItemsEvents_ItemAdded;
            projectItemsEvents.ItemRemoved += ProjectItemsEvents_ItemRemoved;
        }

        private void BindWindowEvents(Events events)
        {
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
        }

        private void RegisterProjectEvents()
        {
            //Manually register the OnAfterOpenProject event on the existing projects 
            //as they are already opened by the time the event would normally be registered
            foreach (ProjectListItem project in Projects)
            {
                if (_vsSolution.GetProjectOfUniqueName(project.Project.UniqueName, out IVsHierarchy projectHierarchy) != VSConstants.S_OK)
                    continue;

                _vsSolutionEvents.OnAfterOpenProject(projectHierarchy, 1);
            }
        }

        #endregion

        #region Events
        public void ProjectItemMoveDeleted(ProjectItem projectItem)
        {
            /*Web site projects are not triggering the same project item added/removed events
            but rather are triggering the hierarchy added/removed events instead - calling the
            existing events here fixes the issue. Moves first do a delete then an add*/
            Project p = projectItem.ContainingProject;
            if (p.Kind.ToUpper() != ExtensionConstants.VsProjectTypeWebSite)
                return;

            ProjectItemsEvents_ItemRemoved(projectItem);

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null)
                return;

            MovedProjectItem movedProjectItem = new MovedProjectItem
            {
                ProjectItem = projectItem,
                OldName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]),
                ProjectItemId = ProjectItemWorker.GetProjectItemId(_vsSolution, SelectedProject.UniqueName, projectItem)
            };

            _movedProjectItems.Add(movedProjectItem);
        }
        public void ProjectItemMoveAdded(ProjectItem projectItem)
        {
            /*Web site projects are not triggering the same project item added/removed events
            but rather are triggering the hierarchy added/removed events instead - calling the
            existing events here fixes the issue. Moves first do a delete then an add*/
            Project p = projectItem.ContainingProject;
            if (p.Kind.ToUpper() != ExtensionConstants.VsProjectTypeWebSite)
                return;

            ProjectItemsEvents_ItemAdded(projectItem);

            if (SelectedProject == null)
                return;

            if (_movedProjectItems.Count == 0)
                return;

            uint addedProjectId = ProjectItemWorker.GetProjectItemId(_vsSolution, SelectedProject.UniqueName, projectItem);

            MovedProjectItem movedProjectItem = _movedProjectItems.FirstOrDefault(m => m.ProjectItemId == addedProjectId);
            if (movedProjectItem == null)
                return;

            ProjectItemEventsOnItemMoved(movedProjectItem.ProjectItem, movedProjectItem.OldName);

            _movedProjectItems.Remove(movedProjectItem);
        }
        private void SolutionEventsOnOpened()
        {
            //ClearConnection();
            //GetProjectsForList();

            //SolutionOpened?.Invoke(this, null);
        }
        private void SolutionEventsOnProjectRemoved(Project project)
        {
            Projects.Remove(Projects.FirstOrDefault(p => p.Project == project));

            OnSolutionProjectRemoved(new SolutionProjectRemovedEventArgs
            {
                Project = project
            });
        }
        private void SolutionEventsOnProjectAdded(Project project)
        {
            foreach (ProjectListItem listProject in Projects)
                if (listProject.Name == project.Name)
                    return;

            Projects.Add(new ProjectListItem
            {
                Project = project,
                Name = project.Name
            });


            OnSolutionProjectAdded(new SolutionProjectAddedEventArgs
            {
                Project = project
            });
        }
        private void SolutionEventsOnProjectRenamed(Project project, string oldName)
        {
            ProjectListItem projectListItem = Projects.FirstOrDefault(p => p.Project == project);
            if (projectListItem != null)
                projectListItem.Name = project.Name;

            OnSolutionProjectRenamed(new SolutionProjectRenamedEventArgs
            {
                Project = project,
                OldName = oldName
            });
        }
        private void SolutionEventsOnBeforeClosing()
        {
            ResetForm();
            ClearConnection();
            ToolWindow = null;
          
            SolutionProjectsList.SelectionChanged -= SolutionProjectsList_OnSelectionChanged;

            SolutionBeforeClosing?.Invoke(this, EventArgs.Empty);
        }
        protected virtual void OnConnected(ConnectEventArgs e)
        {
            var handler = Connected;
            handler?.Invoke(this, e);
        }
        protected virtual void OnSolutionProjectAdded(SolutionProjectAddedEventArgs e)
        {
            var handler = SolutionProjectAdded;
            handler?.Invoke(this, e);
        }
        protected virtual void OnSolutionProjectRemoved(SolutionProjectRemovedEventArgs e)
        {
            var handler = SolutionProjectRemoved;
            handler?.Invoke(this, e);
        }
        protected virtual void OnSolutionProjectRenamed(SolutionProjectRenamedEventArgs e)
        {
            var handler = SolutionProjectRenamed;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectItemRemoved(ProjectItemRemovedEventArgs e)
        {
            var handler = ProjectItemRemoved;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectItemAdded(ProjectItemAddedEventArgs e)
        {
            var handler = ProjectItemAdded;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectItemRenamed(ProjectItemRenamedEventArgs e)
        {
            var handler = ProjectItemRenamed;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectItemMoved(ProjectItemMovedEventArgs e)
        {
            var handler = ProjectItemMoved;
            handler?.Invoke(this, e);
        }
        private void ProjectItemsEvents_ItemRemoved(ProjectItem projectItem)
        {
            OnProjectItemRemoved(new ProjectItemRemovedEventArgs
            {
                ProjectItem = projectItem
            });
        }
        private void ProjectItemsEvents_ItemAdded(ProjectItem projectItem)
        {
            OnProjectItemAdded(new ProjectItemAddedEventArgs
            {
                ProjectItem = projectItem
            });
        }
        private void ProjectItemsEventsOnItemRenamed(ProjectItem projectItem, string oldName)
        {
            OnProjectItemRenamed(new ProjectItemRenamedEventArgs
            {
                ProjectItem = projectItem,
                OldName = oldName
            });
        }
        private void ProjectItemEventsOnItemMoved(ProjectItem postMoveProjectItem, string preMoveName)
        {
            OnProjectItemMoved(new ProjectItemMovedEventArgs
            {
                PreMoveName = preMoveName,
                PostMoveProjectItem = postMoveProjectItem
            });
        }

        #endregion

        private void GetProjectsForList()
        {
            if (ToolWindow != null && ToolWindow.Type == ToolWindowType.PluginTraceViewer)
            {
                SolutionProjectsList.IsEnabled = false;
                return;
            }

            Projects = new ObservableCollection<ProjectListItem>();

            IList<Project> projects = SolutionWorker.GetProjects();
            foreach (Project project in projects)
            {
                Projects.Add(new ProjectListItem
                {
                    Project = project,
                    Name = project.Name
                });
            }

            if (Projects.Any())
            {
                SolutionProjectsList.SelectedIndex = 0;
                SelectedProject = ((ProjectListItem)SolutionProjectsList.SelectedItem).Project;
            }
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusBar.SetStatusBarValue(Resource.StatusBarMessageConnecting, vsStatusAnimation.vsStatusAnimationGeneral);

                CrmLoginForm ctrl = new CrmLoginForm(_autoLogin);
                ctrl.ConnectionToCrmCompleted += ConnectionToCrmCompleted;
                bool? result = ctrl.ShowDialog();

                if (result != true)
                    return;

                if (!ctrl.CrmConnectionMgr?.CrmSvc?.IsReady != true)
                {
                    if (ctrl.CrmConnectionMgr != null)
                        ExceptionHandler.LogCrmConnectionError(Logger, Resource.ErrorMessage_ErrorCrmConnection, ctrl.CrmConnectionMgr);

                    MessageBox.Show(Resource.MessageBox_CannotConnect, Resource.MessageBox_CannotConnect_Title);
                    return;
                }

                OutputLogger.WriteToOutputWindow($@"{Resource.Message_SuccessConnecting}|{ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion}|{ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName}",
                     MessageType.Info);
            }
            finally
            {
                StatusBar.ClearStatusBarValue(vsStatusAnimation.vsStatusAnimationGeneral);
            }
        }

        private void ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (!(sender is CrmLoginForm))
                return;

            CrmLoginForm loginForm = (CrmLoginForm)sender;
            CrmService = loginForm.CrmConnectionMgr.CrmSvc;
            OrganizationId = loginForm.CrmConnectionMgr.ConnectedOrgId;

            if (SolutionProjectsList.SelectedItem != null)
                SelectedProject = ((ProjectListItem)SolutionProjectsList.SelectedItem).Project;

            Dispatcher.Invoke(() =>
            {
                Expander.IsExpanded = false;
                ((CrmLoginForm)sender).Close();
            });

            OnConnected(new ConnectEventArgs
            {
                ServiceClient = loginForm.CrmConnectionMgr.CrmSvc,
            });

            IsConnected = true;

            SetConfigFile();
        }

        private void SolutionProjectsList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox solutionProjectsList = (ComboBox)sender;
            if (solutionProjectsList.SelectedItem == null)
            {
                if (Projects.Count > 0)
                    solutionProjectsList.SelectedItem = Projects[0];
                else
                    return;
            }

            SelectedProject = ((ProjectListItem)solutionProjectsList.SelectedItem).Project;

            Connect.IsEnabled = SelectedProject != null;

            SetConfigFile();

            GetProfiles();

            SelectedProjectChanged?.Invoke(this, e);
        }

        private void ResetForm()
        {
            Projects = new ObservableCollection<ProjectListItem>();
            Profiles = new List<string>();
            SelectedProject = null;
            SelectedProfile = null;
        }

        private void ClearConnection()
        {
            SharedGlobals.SetGlobal("CrmService", null, _dte);
            IsConnected = false;
            CrmService?.Dispose();
            CrmService = null;
        }

        private void AutoLogin_Checked(object sender, RoutedEventArgs e)
        {
            _autoLogin = AutoLogin.ReturnValue();
        }

        private void SetConfigFile()
        {
            if (ToolWindow.Type == ToolWindowType.PluginTraceViewer)
                return;

            if (!Config.ConfigFile.SpklConfigFileExists(ProjectWorker.GetProjectPath(SelectedProject)))
                Config.ConfigFile.CreateSpklConfigFile(SelectedProject);
        }

        private void GetProfiles()
        {
            Profiles = new List<string>();
            Profiles = Config.Profiles.GetProfiles(SelectedProject, ToolWindow.Type);
            if (Profiles == null || Profiles.Count == 0)
            {
                ProfileList.IsEnabled = false;
                return;
            }

            ProfileList.SelectedIndex = 0;

            ProfileList.IsEnabled = Profiles.Count > 1;
        }

        private void ProfileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedProfile = null;

            if (Profiles == null || Profiles.Count == 0)
                return;

            if (ProfileList.SelectedItem != null)
                SelectedProfile = ProfileList.SelectedItem.ToString();

            ProfileChanged?.Invoke(this, e);
        }

        public void CollapsePane()
        {
            Expander.IsExpanded = false;
        }
    }
}