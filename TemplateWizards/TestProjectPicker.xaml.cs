using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;

namespace TemplateWizards
{
    public partial class TestProjectPicker
    {
        private ObservableCollection<ProjectListItem> _projects;
        private ObservableCollection<MockingFrameworkListItem> _mockingFrameworks;

        public MockingFramework SelectedUnitTestFramework { get; set; }
        public Project SelectedProject { get; set; }
        public ObservableCollection<ProjectListItem> Projects
        {
            get => _projects;
            set
            {
                if (value != null && _projects == value)
                    return;

                _projects = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<MockingFrameworkListItem> MockingFrameworks
        {
            get => _mockingFrameworks;
            set
            {
                if (value != null && _mockingFrameworks == value)
                    return;

                _mockingFrameworks = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public TestProjectPicker()
        {
            InitializeComponent();
            DataContext = this;

            GetProjects();
            GetMockingFrameworks();
        }

        private void CreateProject_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void GetMockingFrameworks()
        {
            MockingFrameworks = new ObservableCollection<MockingFrameworkListItem>();
            UnitTestFramework.SelectedIndex = -1;

            string version = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetSdkCoreVersion(SelectedProject);
            Version coreVerion = CrmDeveloperExtensions2.Core.Versioning.StringToVersion(version);

            foreach (MockingFramework mockingFramework in CrmDeveloperExtensions2.Core.Models.MockingFrameworks.GetMockingFrameworks())
            {
                if (mockingFramework.CrmMajorVersion != coreVerion.Major)
                    continue;

                MockingFrameworkListItem mockingFrameworkListItem = new MockingFrameworkListItem
                {
                    Name = mockingFramework.NugetName,
                    MockingFramework = mockingFramework
                };

                MockingFrameworks.Add(mockingFrameworkListItem);
            }
        }

        private void GetProjects()
        {
            Projects = new ObservableCollection<ProjectListItem>();
            IList<Project> projects = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjects(true);

            foreach (Project project in projects)
            {
                ProjectListItem projectListItem = new ProjectListItem
                {
                    Name = project.Name,
                    Project = project
                };

                Projects.Add(projectListItem);
            }

            ProjectToTest.SelectionChanged += ProjectToTest_OnSelectionChanged;
            ProjectToTest.SelectedIndex = 0;
            SelectedProject = Projects[0].Project;
        }

        private void UnitTestFramework_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (UnitTestFramework.SelectedItem == null)
            {
                SelectedUnitTestFramework = null;
                CreateProject.IsEnabled = false;
                return;
            }

            SelectedUnitTestFramework = ((MockingFrameworkListItem)UnitTestFramework.SelectedItem).MockingFramework;
            CreateProject.IsEnabled = true;
        }

        private void ProjectToTest_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ProjectToTest.SelectedItem == null)
                return;

            SelectedProject = ((ProjectListItem)ProjectToTest.SelectedItem).Project;

            GetMockingFrameworks();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}