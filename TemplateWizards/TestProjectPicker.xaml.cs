using System;
using EnvDTE;
using Microsoft.VisualStudio.PlatformUI;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using CrmDeveloperExtensions2.Core.Models;

namespace TemplateWizards
{
    public partial class TestProjectPicker : DialogWindow
    {
        public MockingFramework SelectedUnitTestFramework { get; set; }
        public Project SelectedProject { get; set; }

        public TestProjectPicker()
        {
            InitializeComponent();

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
            List<MockingFramework> mockingFrameworks = MockingFrameworks.Frameworks;
            foreach (MockingFramework mockingFramework in mockingFrameworks)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = mockingFramework.NugetName,
                    Tag = mockingFramework
                };

                UnitTestFramework.Items.Add(item);
            }
        }

        private void GetProjects()
        {
            IList<Project> projects = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetProjects(true);

            foreach (Project project in projects)
            {
                ComboBoxItem item = new ComboBoxItem
                {
                    Content = project.Name,
                    Tag = project
                };

                ProjectToTest.Items.Add(item);
            }
        }

        private void UnitTestFramework_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem unitTestFramework = UnitTestFramework.SelectedItem as ComboBoxItem;
            if (unitTestFramework == null)
            {
                SelectedUnitTestFramework = null;
                return;
            }

            SelectedUnitTestFramework = unitTestFramework.Tag as MockingFramework;
        }

        private void ProjectToTest_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem projectToTest = ProjectToTest.SelectedItem as ComboBoxItem;
            if (projectToTest == null)
                return;

            SelectedProject = projectToTest.Tag as Project;

            string version = CrmDeveloperExtensions2.Core.Vs.ProjectWorker.GetSdkCoreVersion(SelectedProject);
            Version coreVerion = CrmDeveloperExtensions2.Core.Versioning.StringToVersion(version);
            SetEnabledMockingFrameworks(coreVerion);
        }

        private void SetEnabledMockingFrameworks(Version coreVerion)
        {
            if (UnitTestFramework == null)
                return;

            if (coreVerion.Major == 0)
            {
                foreach (var item in UnitTestFramework.Items)
                {
                    ComboBoxItem frameworkItem = (ComboBoxItem)item;
                    frameworkItem.IsEnabled = true;
                }

                return;
            }

            foreach (var item in UnitTestFramework.Items)
            {
                ComboBoxItem frameworkItem = (ComboBoxItem)item;
                MockingFramework framework = frameworkItem.Tag as MockingFramework;
                if (framework != null)
                    frameworkItem.IsEnabled = framework.CrmMajorVersion == coreVerion.Major;
                else
                    frameworkItem.IsEnabled = true;
            }

            if (!((ComboBoxItem)UnitTestFramework.SelectedItem).IsEnabled)
                UnitTestFramework.SelectedIndex = 0;
        }
    }
}