using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Tooling.Connector;
using Window = EnvDTE.Window;

namespace CrmDeveloperExtensions2.Core.Connection
{
    public partial class XrmToolingConnection : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private ObservableCollection<Project> _projects;
        private bool _projectEventsRegistered;
        private readonly IVsSolution _vsSolution;
        private readonly IVsSolutionEvents _vsSolutionEvents;
        private ProjectItem _movedProjectItem;
        private string _movedProjectItemOldName;
        private bool _autoLogin;

        public CrmServiceClient CrmService;
        public Guid OrganizationId;
        public ObservableCollection<Project> Projects
        {
            get => _projects;
            set
            {
                if (_projects == value)
                    return;

                _projects = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Projects"));
            }
        }
        public Project SelectedProject;

        public event EventHandler<ConnectEventArgs> Connected;
        public event EventHandler SolutionBeforeClosing;
        public event EventHandler<SolutionProjectAddedEventArgs> SolutionProjectAdded;
        public event EventHandler<SolutionProjectRemovedEventArgs> SolutionProjectRemoved;
        public event EventHandler<SolutionProjectRenamedEventArgs> SolutionProjectRenamed;
        public event EventHandler<ProjectItemRemovedEventArgs> ProjectItemRemoved;
        public event EventHandler<ProjectItemAddedEventArgs> ProjectItemAdded;
        public event EventHandler<ProjectItemRenamedEventArgs> ProjectItemRenamed;
        public event EventHandler<ProjectItemMovedEventArgs> ProjectItemMoved;
        public event PropertyChangedEventHandler PropertyChanged;

        public XrmToolingConnection()
        {
            InitializeComponent();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            uint solutionEventsCookie;
            _vsSolutionEvents = new VsSolutionEvents(_dte, this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out solutionEventsCookie);

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

            _projects = new ObservableCollection<Project>();
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
        {
            //if (_projects.Count == 0)
            //    GetProjectsForList();

            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //WindowEventsOnWindowActivated in this project can be called when activating another window
            //so we don't want to contine further unless our window is active
            string[] crmDevExWindows = { "A3479AE0-5F4F-4A14-96F4-46F39000023A" };
            if (!crmDevExWindows.Contains(gotFocus.ObjectKind.Replace("{", String.Empty).Replace("}", String.Empty)))
                return;

            if (_projects.Count == 0)
                GetProjectsForList();

            if (!_projectEventsRegistered)
            {
                RegisterProjectEvents();
                _projectEventsRegistered = true;
            }
        }

        private void RegisterProjectEvents()
        {
            //Manually register the OnAfterOpenProject event on the existing projects 
            //as they are already opened by the time the event would normally be registered
            foreach (Project project in _projects)
            {
                IVsHierarchy projectHierarchy;
                if (_vsSolution.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) != VSConstants.S_OK)
                    continue;

                _vsSolutionEvents.OnAfterOpenProject(projectHierarchy, 1);
            }
        }

        public void ProjectItemMoveDeleted(ProjectItem projectItem)
        {
            _movedProjectItem = projectItem;

            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            _movedProjectItemOldName = FileSystem.LocalPathToCrmPath(projectPath, projectItem.FileNames[1]);
        }

        public void ProjectItemMoveAdded(ProjectItem projectItem)
        {
            if (SelectedProject == null)
                return;

            if (_movedProjectItem == null)
                return;

            uint movedProjectId = ProjectItemWorker.GetProjectItemId(_vsSolution, SelectedProject.UniqueName, _movedProjectItem);
            uint addedProjectId = ProjectItemWorker.GetProjectItemId(_vsSolution, SelectedProject.UniqueName, projectItem);

            if (movedProjectId != addedProjectId)
                return;

            ProjectItemEventsOnItemMoved(projectItem, _movedProjectItemOldName);
            _movedProjectItem = null;
        }

        private void SolutionEventsOnOpened()
        {
            GetProjectsForList();
        }
        private void SolutionEventsOnProjectRemoved(Project project)
        {
            _projects.Remove(project);

            OnSolutionProjectRemoved(new SolutionProjectRemovedEventArgs
            {
                Project = project
            });
        }
        private void SolutionEventsOnProjectAdded(Project project)
        {
            foreach (Project listProject in _projects)
                if (listProject.Name == project.Name)
                    return;

            _projects.Add(project);

            OnSolutionProjectAdded(new SolutionProjectAddedEventArgs
            {
                Project = project
            });
        }
        private void SolutionEventsOnProjectRenamed(Project project, string oldName)
        {
            Project selectedProject = (Project)SolutionProjectsList.SelectedItem;
            SolutionProjectsList.ItemsSource = null;
            SolutionProjectsList.ItemsSource = _projects;
            SolutionProjectsList.SelectedItem = selectedProject;
            OutputLogger.DeleteOutputWindow();
            OnSolutionProjectRenamed(new SolutionProjectRenamedEventArgs
            {
                Project = project,
                OldName = oldName
            });
        }
        private void SolutionEventsOnBeforeClosing()
        {
            ResetForm();

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
            IList<Project> projects = SolutionWorker.GetProjects();
            foreach (Project project in projects)
            {
                _projects.Add(project);
            }

            SolutionProjectsList.ItemsSource = _projects;
            SolutionProjectsList.DisplayMemberPath = "Name";

            if (_projects.Any())
                SolutionProjectsList.SelectedIndex = 0;
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusBar.SetStatusBarValue(_dte, Core.Resources.Resource.StatusBarMessageConnecting, vsStatusAnimation.vsStatusAnimationGeneral);

                CrmLoginForm ctrl = new CrmLoginForm(_autoLogin);
                ctrl.ConnectionToCrmCompleted += ConnectionToCrmCompleted;
                bool? result = ctrl.ShowDialog();

                if (result != true)
                    return;

                if (ctrl.CrmConnectionMgr?.CrmSvc == null || !ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                {
                    if (ctrl.CrmConnectionMgr != null)
                        OutputLogger.WriteToOutputWindow("Error connecting to CRM: Last error: " +
                                                         ctrl.CrmConnectionMgr.LastError + Environment.NewLine +
                                                         "Last exception: " +
                                                         ctrl.CrmConnectionMgr.LastException.Message +
                                                         Environment.NewLine +
                                                         ctrl.CrmConnectionMgr.LastException.StackTrace,
                            MessageType.Error);
                    MessageBox.Show("Cannot connect to CRM - see Output window for details", "Connection Status");
                    return;
                }

                OutputLogger.WriteToOutputWindow("Connected to CRM! Version: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion +
                    " Org: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, MessageType.Info);
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

            Dispatcher.Invoke(() =>
            {
                Expander.IsExpanded = false;
                ((CrmLoginForm)sender).Close();
            });

            OnConnected(new ConnectEventArgs
            {
                ServiceClient = loginForm.CrmConnectionMgr.CrmSvc
            });
        }

        private void SolutionProjectsList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox solutionProjectsList = (ComboBox)sender;
            SelectedProject = (Project)solutionProjectsList.SelectedItem;
            Connect.IsEnabled = SelectedProject != null;
        }

        private void SolutionProjectsList_Initialized(object sender, EventArgs e)
        {
        }

        private void ResetForm()
        {
            _projects = new ObservableCollection<Project>();
            SolutionProjectsList.ItemsSource = null;
        }

        private void AutoLogin_Checked(object sender, RoutedEventArgs e)
        {
            _autoLogin = AutoLogin.IsChecked.HasValue && AutoLogin.IsChecked.Value;
        }
    }
}
