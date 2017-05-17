using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using CrmDeveloperExtensions.Core.Models;
using CrmDeveloperExtensions.Core.Vs;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Tooling.Connector;
using Window = EnvDTE.Window;

namespace CrmDeveloperExtensions.Core.Connection
{
    public partial class XrmToolingConnection : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private ObservableCollection<Project> _projects;
        private bool _projectEventsRegistered;
        private IVsSolution _vsSolution;

        public CrmServiceClient CrmService;

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

        public event EventHandler<ConnectEventArgs> Connected;
        public event EventHandler<ProjectAddedEventArgs> ProjectAdded;
        public event EventHandler<ProjectRemovedEventArgs> ProjectRemoved;
        public event EventHandler<ProjectRenamedEventArgs> ProjectRenamed;
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
            //if (_projects == null)
            //    _projects = GetProjectsForList();

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
            OnProjectRemoved(new ProjectRemovedEventArgs
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
            OnProjectAdded(new ProjectAddedEventArgs
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

            OnProjectRenamed(new ProjectRenamedEventArgs
            {
                Project = project,
                OldName = oldName
            });
        }
        private void SolutionEventsOnBeforeClosing()
        {
            ResetForm();
        }
        protected virtual void OnConnected(ConnectEventArgs e)
        {
            var handler = Connected;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectAdded(ProjectAddedEventArgs e)
        {
            var handler = ProjectAdded;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectRemoved(ProjectRemovedEventArgs e)
        {
            var handler = ProjectRemoved;
            handler?.Invoke(this, e);
        }
        protected virtual void OnProjectRenamed(ProjectRenamedEventArgs e)
        {
            var handler = ProjectRenamed;
            handler?.Invoke(this, e);
        }

        public void ProjectItemAdded()
        {
            MessageBox.Show("Added");
        }

        private void GetProjectsForList()
        {
            IList<Project> projects = SolutionWorker.GetProjects();
            foreach (Project project in projects)
            {
                if (!ProjectWorker.IsProjectLoaded(project))
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
                    MessageBox.Show("Connected to CRM! Version: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion +
                                    " Org: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, "Connection Status");
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

            OnConnected(new ConnectEventArgs
            {
                ServiceClient = loginForm.CrmConnectionMgr.CrmSvc
            });

            Dispatcher.Invoke(() =>
            {
                ((CrmLoginForm)sender).Close();
            });
        }

        private void SolutionProjectsList_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {

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
