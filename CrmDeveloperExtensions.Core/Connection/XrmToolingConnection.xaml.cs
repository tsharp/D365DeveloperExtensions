using CrmDeveloperExtensions.Core.Vs;
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
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Window = EnvDTE.Window;

namespace CrmDeveloperExtensions.Core.Connection
{
    public partial class XrmToolingConnection : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private ObservableCollection<Project> _projects;
        private bool _projectEventsRegistered;
        private readonly IVsSolution _vsSolution;

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
            IVsSolutionEvents vsSolutionEvents = new VsSolutionEvents(_dte, this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(vsSolutionEvents, out solutionEventsCookie);

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

            if (!_projectEventsRegistered)
            {
                RegisterProjectEvents();
                _projectEventsRegistered = true;
            }

            GetProjectsForList();

            //if (!gotFocus.Caption.StartsWith("CRM DevEx")) return;

            //ProjectsDdl.IsEnabled = true;
            //AddConnection.IsEnabled = true;
            //Connections.IsEnabled = true;

            //foreach (var project in GetProjects())
            //{
            //    SolutionProjectAdded(project);
            //}
        }

        private void RegisterProjectEvents()
        {
            //Manually register the OnAfterOpenProject event on the existing projects as they are already opened by the time the event would normally be registered
            foreach (Project project in _projects)
            {
                IVsHierarchy projectHierarchy;
                if (_vsSolution.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) != VSConstants.S_OK)
                    continue;

                IVsSolutionEvents vsSolutionEvents = new VsSolutionEvents(_dte, this);
                vsSolutionEvents.OnAfterOpenProject(projectHierarchy, 1);
            }
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
            {
                if (listProject.Name == project.Name)
                    return;
            }

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

        private void GetProjectsForList()
        {
            IList<Project> projects = SolutionWorker.GetProjects();
            foreach (Project project in projects)
            {
                //if (!ProjectWorker.IsProjectLoaded(project))
                _projects.Add(project);
            }

            SolutionProjectsList.ItemsSource = _projects;
            SolutionProjectsList.DisplayMemberPath = "Name";
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusBar.SetStatusBarValue(_dte, Core.Resources.Resource.StatusBarMessageConnecting, vsStatusAnimation.vsStatusAnimationGeneral);

                CrmLoginForm ctrl = new CrmLoginForm();
                ctrl.ConnectionToCrmCompleted += ConnectionToCrmCompleted;
                ctrl.ShowDialog();

                if (ctrl.CrmConnectionMgr?.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                {
                    //MessageBox.Show("Connected to CRM! Version: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion +
                    //                " Org: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, "Connection Status");
                }
                else
                {
                    MessageBox.Show("Cannot connect; try again!", "Connection Status");
                }
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
        }

        private void SolutionProjectsList_Initialized(object sender, EventArgs e)
        {
        }

        private void ResetForm()
        {
            _projects = new ObservableCollection<Project>();
            SolutionProjectsList.ItemsSource = null;
        }
    }
}
