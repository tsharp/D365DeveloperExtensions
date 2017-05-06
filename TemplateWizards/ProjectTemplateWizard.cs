using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using System.Xml.Linq;
using CrmDeveloperExtensions.Core;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using NuGet.VisualStudio;


namespace TemplateWizards
{
    public class ProjectTemplateWizard : IWizard
    {
        private DTE _dte;
        private string _coreVersion;
        private string _workflowVersion;
        private string _clientVersion;
        private string _clientPackage;
        private string _project;
        private string _crmProjectType = "Plug-in";
        private bool _needsCore;
        private bool _needsWorkflow;
        private bool _needsClient;
        private bool _isUnitTest;
        private bool _isUnitTestItem;
        private bool _signAssembly;
        private bool _isNunit;
        private string _destDirectory;

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            try
            {
                _dte = (DTE)automationObject;

                if (replacementsDictionary.ContainsKey("$destinationdirectory$"))
                    _destDirectory = replacementsDictionary["$destinationdirectory$"];

                if (replacementsDictionary.ContainsKey("$wizarddata$"))
                {
                    string wizardData = replacementsDictionary["$wizarddata$"];
                    ReadWizardData(wizardData);
                }

                //Display the form prompting for the SDK version and/or project to unit test against
                if (_needsCore)
                {
                    var sdkVersionPicker = new SdkVersionPicker(_needsWorkflow, _needsClient)
                    {
                        HasMinimizeButton = false,
                        HasMaximizeButton = false,
                        ResizeMode = ResizeMode.NoResize
                    };
                    bool? result = sdkVersionPicker.ShowModal();

                    if (!result.HasValue || result.Value == false)
                        throw new WizardBackoutException();

                    _coreVersion = sdkVersionPicker.CoreVersion;
                    _workflowVersion = sdkVersionPicker.WorkflowVersion;
                    _clientVersion = sdkVersionPicker.ClientVersion;
                    _clientPackage = sdkVersionPicker.ClientPackage;
                }

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
                MessageBox.Show(string.Format("Error occurred running wizard:\n\n{0}", ex));
                throw new WizardCancelledException("Internal error", ex);
            }
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null) return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            switch (_crmProjectType)
            {
                case "Console":
                case "Plug-in":
                case "Workflow":
                    HandleCrmAssemblyProjects(project, installer);
                    break;
                case "WebResource":
                    HandleWebResourceProjects(project);
                    break;
                    //case "TypeScript":
                    //    HandleTypeScriptProject(project, installer);
                    //    break;
                    //case "Package":
                    //    HandleSolutionPackagerProject(project);
                    //    break;
            }
        }

        private void HandleWebResourceProjects(Project project)
        {
            //Turn off Build option in build configurations
            SolutionConfigurations solutionConfigurations = _dte.Solution.SolutionBuild.SolutionConfigurations;
            string folderProjectFileName = ProjectWorker.GetFolderProjectFileName(project.FullName);
            SolutionWorker.SetBuildConfigurationOff(solutionConfigurations, folderProjectFileName);
        }

        private void HandleCrmAssemblyProjects(Project project, IVsPackageInstaller installer)
        {
            try
            {
                project.DTE.SuppressUI = true;

                //Pre-2015 use .NET 4.0
                if (Versioning.StringToVersion(_coreVersion).Major < 7)
                    project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.0";


                //Install all the NuGet packages
                project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                NuGetProcessor.InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", _coreVersion);
                if (_needsWorkflow)
                    NuGetProcessor.InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", _coreVersion);
                if (_needsClient)
                    NuGetProcessor.InstallPackage(installer, project, _clientPackage, _clientVersion);


                ProjectWorker.ExcludeFolder(project, "bin");

                if (_signAssembly)
                    Signing.GenerateKey(_dte, project, _destDirectory);

                if (_isUnitTest)
                    return;

                //InstallPackage(installer, project, "Moq", "4.5.28");
                //if (_isNunit)
                //{
                //    InstallPackage(installer, project, "NUnitTestAdapter.WithFramework", "2.0.0");
                //    AddSetting(project, "CRMTestType", "NUNIT");
                //}
                //else
                //    AddSetting(project, "CRMTestType", "UNIT");

                ProjectWorker.ExcludeFolder(project, "performance");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: " + ex.Message);
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
                                    _crmProjectType = reader.Value;
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
                                case "IsUnitTestItem":
                                    _isUnitTestItem = bool.Parse(reader.Value);
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
