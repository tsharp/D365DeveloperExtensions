using CrmDeveloperExtensions2.Core.Enums;
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
        private readonly DTE _dte;
        private readonly Solution _solution;
        private ObservableCollection<ProjectListItem> _projects;
        private List<string> _profiles;
        private bool _projectEventsRegistered;
        private readonly IVsSolution _vsSolution;
        private readonly IVsSolutionEvents _vsSolutionEvents;
        private readonly List<MovedProjectItem> _movedProjectItems;
        private bool _autoLogin;
        private bool _isConnected;
        private Project _selectedProject;
        private string _selectedProfile;

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

            _vsSolutionEvents = new VsSolutionEvents(_dte, this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out uint solutionEventsCookie);

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;

            var events2 = (Events2)_dte.Events;
            var projectItemsEvents = events2.ProjectItemsEvents;
            projectItemsEvents.ItemRenamed += ProjectItemsEventsOnItemRenamed;
            projectItemsEvents.ItemAdded += ProjectItemsEvents_ItemAdded;
            projectItemsEvents.ItemRemoved += ProjectItemsEvents_ItemRemoved;

            var solutionEvents = events.SolutionEvents;
            solutionEvents.ProjectAdded += SolutionEventsOnProjectAdded;
            solutionEvents.ProjectRemoved += SolutionEventsOnProjectRemoved;
            solutionEvents.ProjectRenamed += SolutionEventsOnProjectRenamed;
            solutionEvents.BeforeClosing += SolutionEventsOnBeforeClosing;
            solutionEvents.Opened += SolutionEventsOnOpened;

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

            if (Projects == null || Projects.Count == 0)
                GetProjectsForList();

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
        public void ProjectItemMoveDeleted(ProjectItem projectItem)
        {
            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

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
            ClearConnection();
            GetProjectsForList();

            SolutionOpened?.Invoke(this, null);
        }
        private void SolutionEventsOnProjectRemoved(Project project)
        {
            Projects.Remove(Projects.FirstOrDefault(p => p.Project == project));
            Profiles = new List<string>();

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
        private void GetProjectsForList()
        {
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
                SolutionProjectsList.SelectedIndex = 0;
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusBar.SetStatusBarValue(_dte, Resource.StatusBarMessageConnecting, vsStatusAnimation.vsStatusAnimationGeneral);

                CrmLoginForm ctrl = new CrmLoginForm(_autoLogin);
                ctrl.ConnectionToCrmCompleted += ConnectionToCrmCompleted;
                bool? result = ctrl.ShowDialog();

                if (result != true)
                    return;

                if (ctrl.CrmConnectionMgr?.CrmSvc == null || !ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                {
                    if (ctrl.CrmConnectionMgr != null)
                        OutputLogger.WriteToOutputWindow(Resource.OutputLogger_ErrorConnecting_Error + ": " +
                                                         ctrl.CrmConnectionMgr.LastError + Environment.NewLine +
                                                         Resource.OutputLogger_ErrorConnecting_Exception + ": " +
                                                         ctrl.CrmConnectionMgr.LastException.Message +
                                                         Environment.NewLine +
                                                         ctrl.CrmConnectionMgr.LastException.StackTrace,
                            MessageType.Error);
                    MessageBox.Show(Resource.MessageBox_CannotConnect, Resource.MessageBox_CannotConnect_Title);
                    return;
                }

                OutputLogger.WriteToOutputWindow(Resource.OutputLogger_SuccessConnecting + " | " +
                    ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion + " | " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, MessageType.Info);
            }
            finally
            {
                StatusBar.ClearStatusBarValue(_dte, vsStatusAnimation.vsStatusAnimationGeneral);
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

            SharedGlobals.SetGlobal("CrmService", loginForm.CrmConnectionMgr.CrmSvc, _dte);

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
                return;

            SelectedProject = ((ProjectListItem)solutionProjectsList.SelectedItem).Project;

            Connect.IsEnabled = SelectedProject != null;

            SetConfigFile();

            GetProfiles();

            SelectedProjectChanged?.Invoke(this, e);
        }

        private void SolutionProjectsList_Initialized(object sender, EventArgs e)
        {
        }

        private void ResetForm()
        {
            Projects = new ObservableCollection<ProjectListItem>();
        }

        private void ClearConnection()
        {
            SharedGlobals.SetGlobal("CrmService", null, _dte);
            IsConnected = false;
            CrmService?.Dispose();
            CrmService = null;
            Projects = null;
            Profiles = null;
            SelectedProject = null;
        }

        private void AutoLogin_Checked(object sender, RoutedEventArgs e)
        {
            _autoLogin = AutoLogin.IsChecked.HasValue && AutoLogin.IsChecked.Value;
        }

        private void SetConfigFile()
        {
            if (!Config.ConfigFile.SpklConfigFileExists(ProjectWorker.GetProjectPath(SelectedProject)))
                Config.ConfigFile.CreateSpklConfigFile(SelectedProject);
        }

        private void GetProfiles()
        {
            ToolWindow toolWindow = HostWindow.GetCrmDevExWindow(_dte.ActiveWindow);
            Profiles = Config.Profiles.GetProfiles(ProjectWorker.GetProjectPath(SelectedProject), toolWindow.Type);
            if (Profiles == null || Profiles.Count == 0)
            {
                ProfileList.IsEnabled = false;
                return;
            }

            ProfileList.SelectedIndex = 0;

            if (Profiles.Count == 1)
                ProfileList.IsEnabled = false;
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
    }
}