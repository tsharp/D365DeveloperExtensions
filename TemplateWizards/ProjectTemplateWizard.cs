using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.UserOptions;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using NLog;
using NuGet.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml;
using TemplateWizards.Models;
using TemplateWizards.Resources;
using WizardCancelledException = Microsoft.VisualStudio.TemplateWizard.WizardCancelledException;

namespace TemplateWizards
{
    public class ProjectTemplateWizard : IWizard
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private DTE _dte;
        private string _coreVersion;
        private string _clientVersion;
        private string _clientPackage;
        private ProjectType _crmProjectType = ProjectType.Plugin;
        private bool _needsCore;
        private bool _needsWorkflow;
        private bool _needsClient;
        private bool _isUnitTest;
        private bool _signAssembly;
        private string _destDirectory;
        private string _unitTestFrameworkPackage;
        private CustomTemplate _customTemplate;
        private bool _addFile = true;
        private string _typesXrmVersion;

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            try
            {
                _dte = (DTE)automationObject;

                ProjectDataHandler.AddOrUpdateReplacements("$referenceproject$", "False", ref replacementsDictionary);
                if (replacementsDictionary.ContainsKey("$destinationdirectory$"))
                    _destDirectory = replacementsDictionary["$destinationdirectory$"];

                if (replacementsDictionary.ContainsKey("$wizarddata$"))
                {
                    string wizardData = replacementsDictionary["$wizarddata$"];
                    ReadWizardData(wizardData);
                }

                if (_isUnitTest)
                    PreHandleUnitTestProjects(replacementsDictionary);

                if (_needsCore)
                    PreHandleCrmAssemblyProjects(replacementsDictionary);

                if (_crmProjectType == ProjectType.CustomItem)
                    replacementsDictionary = PreHandleCustomItem(replacementsDictionary);

                if (_crmProjectType == ProjectType.TypeScript)
                    PreHandleTypeScriptProjects();
            }
            catch (WizardBackoutException)
            {
                try
                {
                    DirectoryInfo destination = new DirectoryInfo(replacementsDictionary["$destinationdirectory$"]);
                    Directory.Delete(replacementsDictionary["$destinationdirectory$"]);
                    //Delete solution directory if empty
                    if (destination.Parent != null && FileSystem.IsDirectoryEmpty(replacementsDictionary["$solutiondirectory$"]))
                        Directory.Delete(replacementsDictionary["$solutiondirectory$"]);
                }
                catch
                {
                    // If it fails (doesn't exist/contains files/read-only), let the directory stay.
                }
                throw;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_TemplateWizardError, ex);
                MessageBox.Show(Resource.ErrorMessage_TemplateWizardError);
                throw new WizardCancelledException("Internal error", ex);
            }
        }

        private Dictionary<string, string> PreHandleCustomItem(Dictionary<string, string> replacementsDictionary)
        {
            string templateFolder = UserOptionsHelper.GetOption<string>(UserOptionProperties.CustomTemplatesPath);
            _addFile = CustomTemplateHandler.ValidateTemplateFolder(templateFolder);
            if (!_addFile)
                return replacementsDictionary;

            _addFile = CustomTemplateHandler.ValidateTemplateFile(templateFolder);
            if (!_addFile)
                return replacementsDictionary;

            CustomTemplates templates = CustomTemplateHandler.GetTemplateConfig(templateFolder);
            if (templates == null)
            {
                _addFile = false;
                return replacementsDictionary;
            }

            List<CustomTemplate> results = CustomTemplateHandler.GetTemplatesByLanguage(templates, "CSharp");
            if (results.Count == 0)
            {
                MessageBox.Show(Resource.MessageBox_AddCustomTemplate);
                _addFile = false;
                return replacementsDictionary;
            }

            CustomTemplatePicker templatePicker = CustomTemplateHandler.GetCustomTemplate(results);
            if (templatePicker.SelectedTemplate == null)
            {
                _addFile = false;
                return replacementsDictionary;
            }

            _customTemplate = templatePicker.SelectedTemplate;

            string content = CustomTemplateHandler.GetTemplateContent(templateFolder, _customTemplate, replacementsDictionary);

            replacementsDictionary.Add("$customtemplate$", content);

            return replacementsDictionary;
        }

        private void PreHandleCrmAssemblyProjects(Dictionary<string, string> replacementsDictionary)
        {
            var sdkVersionPicker = new SdkVersionPicker(_needsWorkflow, _needsClient);
            bool? result = sdkVersionPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            _coreVersion = sdkVersionPicker.CoreVersion;
            _clientVersion = sdkVersionPicker.ClientVersion;
            _clientPackage = sdkVersionPicker.ClientPackage;

            if (!string.IsNullOrEmpty(_clientVersion))
            {
                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    Versioning.StringToVersion(_clientVersion).Major >= 8 ? "1" : "0", ref replacementsDictionary);
            }
        }

        private void PreHandleUnitTestProjects(Dictionary<string, string> replacementsDictionary)
        {
            var testProjectPicker = new TestProjectPicker();
            bool? result = testProjectPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            if (testProjectPicker.SelectedProject != null)
            {
                Solution solution = _dte.Solution;
                Project project = testProjectPicker.SelectedProject;
                string path = string.Empty;
                string projectPath = Path.GetDirectoryName(project.FullName);
                string solutionPath = Path.GetDirectoryName(solution.FullName);
                if (!string.IsNullOrEmpty(projectPath) && !string.IsNullOrEmpty(solutionPath))
                {
                    if (projectPath.StartsWith(solutionPath))
                        path = "..\\" + project.UniqueName;
                    else
                        path = project.FullName;
                }

                ProjectDataHandler.AddOrUpdateReplacements("$referenceproject$", "True", ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectPath$", path, ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectId$", project.Kind, ref replacementsDictionary);
                ProjectDataHandler.AddOrUpdateReplacements("$projectName$", project.Name, ref replacementsDictionary);
            }

            if (testProjectPicker.SelectedUnitTestFramework != null)
            {
                _unitTestFrameworkPackage = testProjectPicker.SelectedUnitTestFramework.NugetName;

                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    testProjectPicker.SelectedUnitTestFramework.CrmMajorVersion >= 8 ? "1" : "0", ref replacementsDictionary);
            }
            else
            {
                if (testProjectPicker.SelectedProject == null)
                    return;

                string version = ProjectWorker.GetSdkCoreVersion(testProjectPicker.SelectedProject);
                ProjectDataHandler.AddOrUpdateReplacements("$useXrmToolingClientUsing$",
                    Versioning.StringToVersion(version).Major >= 8 ? "1" : "0", ref replacementsDictionary);
            }
        }

        private void PreHandleTypeScriptProjects()
        {
            NpmHistory history = NpmProcessor.GetPackageHistory("@types/xrm");

            NpmPicker npmPicker = new NpmPicker(history);
            bool? result = npmPicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            _typesXrmVersion = npmPicker.SelectedPackage.Version;
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return _addFile;
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
            if (!_addFile)
                return;

            if (_crmProjectType != ProjectType.CustomItem)
                return;

            if (_customTemplate == null)
                return;

            Project project = projectItem.ContainingProject;

            CustomTemplateHandler.AddTemplateReferences(_customTemplate, project);

            CustomTemplateHandler.InstallTemplateNuGetPackages(_dte, _customTemplate, project);

            if (!string.IsNullOrEmpty(_customTemplate.FileName))
                projectItem.Name = _customTemplate.FileName;
        }

        public void ProjectFinishedGenerating(Project project)
        {
            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null)
                return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            switch (_crmProjectType)
            {
                case ProjectType.UnitTest:
                    PostHandleUnitTestProjects(project, installer);
                    break;
                case ProjectType.Console:
                case ProjectType.Plugin:
                case ProjectType.Workflow:
                    PostHandleCrmAssemblyProjects(project, installer);
                    break;
                case ProjectType.TypeScript:
                    PostHandleTypeScriptProject(project);
                    break;
                case ProjectType.SolutionPackage:
                    PostHandleSolutionPackagerProject(project);
                    break;
            }

            if (_isUnitTest)
                PostHandleUnitTestProjects(project, installer);

            _dte.ExecuteCommand("File.SaveAll");
        }

        private void PostHandleSolutionPackagerProject(Project project)
        {
            foreach (SolutionConfiguration solutionConfiguration in _dte.Solution.SolutionBuild.SolutionConfigurations)
                foreach (SolutionContext solutionContext in solutionConfiguration.SolutionContexts)
                    solutionContext.ShouldBuild = false;

            //Delete bin & obj folders
            Directory.Delete(Path.GetDirectoryName(project.FullName) + "//bin", true);
            Directory.Delete(Path.GetDirectoryName(project.FullName) + "//obj", true);

            project.ProjectItems.AddFolder("package");
        }

        private void PostHandleTypeScriptProject(Project project)
        {
            NpmProcessor.InstallPackage("@types/xrm", _typesXrmVersion, ProjectWorker.GetProjectPath(project));

            _dte.ExecuteCommand("ProjectandSolutionContextMenus.CrossProjectMultiItem.RefreshFolder");
        }

        private void PostHandleUnitTestProjects(Project project, IVsPackageInstaller installer)
        {
            NuGetProcessor.InstallPackage(installer, project, ExtensionConstants.MsTestTestAdapter, null);
            NuGetProcessor.InstallPackage(installer, project, ExtensionConstants.MsTestTestFramework, null);

            if (_unitTestFrameworkPackage != null)
                NuGetProcessor.InstallPackage(installer, project, _unitTestFrameworkPackage, null);
        }

        private void PostHandleCrmAssemblyProjects(Project project, IVsPackageInstaller installer)
        {
            try
            {
                project.DTE.SuppressUI = true;

                //Pre-2015 use .NET 4.0
                if (Versioning.StringToVersion(_coreVersion).Major < 7)
                    project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.0";

                //Install all the NuGet packages
                project = (Project)((Array)_dte.ActiveSolutionProjects).GetValue(0);
                NuGetProcessor.InstallPackage(installer, project, Resource.SdkAssemblyCore, _coreVersion);
                if (_needsWorkflow)
                    NuGetProcessor.InstallPackage(installer, project, Resource.SdkAssemblyWorkflow, _coreVersion);
                if (_needsClient)
                    NuGetProcessor.InstallPackage(installer, project, _clientPackage, _clientVersion);

                ProjectWorker.ExcludeFolder(project, "bin");
                ProjectWorker.ExcludeFolder(project, "performance");

                if (_signAssembly)
                    Signing.GenerateKey(project, _destDirectory);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorProcessingTemplate, ex);
                MessageBox.Show(Resource.ErrorMessage_ErrorProcessingTemplate);
            }
        }

        private void ReadWizardData(string wizardData)
        {
            XmlReaderSettings settings = new XmlReaderSettings { ConformanceLevel = ConformanceLevel.Fragment };

            string el = "";

            using (XmlReader reader = XmlReader.Create(new StringReader(wizardData), settings))
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            el = reader.Name;
                            break;
                        case XmlNodeType.Text:
                            switch (el)
                            {
                                case "CRMProjectType":
                                    _crmProjectType = (ProjectType)Enum.Parse(typeof(ProjectType), reader.Value);
                                    break;
                                case "NeedsCore":
                                    _needsCore = bool.Parse(reader.Value);
                                    break;
                                case "NeedsWorkflow":
                                    _needsWorkflow = bool.Parse(reader.Value);
                                    break;
                                case "NeedsClient":
                                    _needsClient = bool.Parse(reader.Value);
                                    break;
                                case "IsUnitTest":
                                    _isUnitTest = bool.Parse(reader.Value);
                                    break;
                                case "SignAssembly":
                                    _signAssembly = bool.Parse(reader.Value);
                                    break;
                            }
                            break;
                    }
                }
        }
    }
}