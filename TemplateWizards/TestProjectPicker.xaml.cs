using D365DeveloperExtensions.Core.Models;
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

            string version = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetSdkCoreVersion(SelectedProject);
            Version coreVerion = D365DeveloperExtensions.Core.Versioning.StringToVersion(version);

            foreach (MockingFramework mockingFramework in D365DeveloperExtensions.Core.Models.MockingFrameworks.GetMockingFrameworks())
            {
                if (mockingFramework.CrmMajorVersion != coreVerion.Major)
                    continue;

                var mockingFrameworkListItem = CreateMockingFrameworkItem(mockingFramework);

                MockingFrameworks.Add(mockingFrameworkListItem);
            }
        }

        private static MockingFrameworkListItem CreateMockingFrameworkItem(MockingFramework mockingFramework)
        {
            MockingFrameworkListItem mockingFrameworkListItem = new MockingFrameworkListItem
            {
                Name = mockingFramework.NugetName,
                MockingFramework = mockingFramework
            };

            return mockingFrameworkListItem;
        }

        private void GetProjects()
        {
            Projects = new ObservableCollection<ProjectListItem>();
            IList<Project> projects = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjects(true);

            foreach (Project project in projects)
            {
                var projectListItem = CreateProjectItem(project);

                Projects.Add(projectListItem);
            }

            ProjectToTest.SelectionChanged += ProjectToTest_OnSelectionChanged;
            ProjectToTest.SelectedIndex = 0;
            SelectedProject = Projects[0].Project;
        }

        private static ProjectListItem CreateProjectItem(Project project)
        {
            ProjectListItem projectListItem = new ProjectListItem
            {
                Name = project.Name,
                Project = project
            };

            return projectListItem;
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