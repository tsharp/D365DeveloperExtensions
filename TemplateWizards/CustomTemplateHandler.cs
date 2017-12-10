using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using Newtonsoft.Json;
using NuGet.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Resources;
using TemplateWizards.Models;
using VSLangProj;

namespace TemplateWizards
{
    public class CustomTemplateHandler
    {
        public static CustomTemplates GetTemplateConfig(string templateFolder)
        {
            string path = Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile);
            if (!File.Exists(path))
                return null;

            try
            {
                CustomTemplates templates;
                using (StreamReader file = File.OpenText(path))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    templates = (CustomTemplates)serializer.Deserialize(file, typeof(CustomTemplates));
                }

                return templates;
            }
            catch
            {
                throw new Exception($"Unable to read or deserialize {path}");
            }
        }

        public static List<CustomTemplate> GetTemplatesByLanguage(CustomTemplates templates, string language)
        {
            return templates.Templates
                .Where(t => t.Language.Equals(language, StringComparison.InvariantCultureIgnoreCase)).ToList();
        }

        public static string GetTemplateContent(string templateFolder, CustomTemplate template, Dictionary<string, string> replacementsDictionary)
        {
            string path = Path.Combine(templateFolder, template.RelativePath);
            if (!File.Exists(path))
                return null;

            string content = File.ReadAllText(path);

            if (!template.CoreReplacements)
                return content;

            foreach (KeyValuePair<string, string> keyValuePair in replacementsDictionary)
                content = content.Replace(keyValuePair.Key, keyValuePair.Value);

            return content;
        }

        public static void InstallTemplateNuGetPackages(DTE dte, CustomTemplate customTemplate, Project project)
        {
            if (customTemplate.CustomTemplateNuGetPackages.Count <= 0)
                return;

            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null)
                return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            foreach (CustomTemplateNuGetPackage package in customTemplate.CustomTemplateNuGetPackages)
                AddPackage(dte, installer, package, project);
        }

        private static void AddPackage(DTE dte, IVsPackageInstaller installer, CustomTemplateNuGetPackage package, Project project)
        {
            string packageVersion = string.IsNullOrEmpty(package.Version) ?
                null :
                package.Version;

            try
            {
                NuGetProcessor.InstallPackage(installer, project, package.Name, packageVersion);
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    $"Failed to add NuGet package {package.Name} {packageVersion}: {Environment.NewLine}{ex.Message}{Environment.NewLine}{ex.StackTrace}", MessageType.Error);
            }
        }

        public static void AddTemplateReferences(CustomTemplate customTemplate, Project project)
        {
            if (customTemplate.CustomTemplateReferences.Count <= 0)
                return;

            if (!(project?.Object is VSProject vsproject))
                return;

            foreach (CustomTemplateReference reference in customTemplate.CustomTemplateReferences)
                CrmDeveloperExtensions2.Core.Vs.ProjectWorker.AddProjectReference(vsproject, reference.Name);
        }

        public static string GetTemplateFileTemplate()
        {
            StreamResourceInfo streamResourceInfo = Application.GetResourceStream(new Uri("TemplateWizards;component/Template/templates.json",
                UriKind.Relative));

            if (streamResourceInfo == null)
                return null;

            StreamReader sr = new StreamReader(streamResourceInfo.Stream);
            return sr.ReadToEnd();
        }

        public static void CreateTemplateFileTemplate(string templateFolder)
        {
            string content = GetTemplateFileTemplate();
            string path = Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile);

            FileStream file = File.Create(path);
            file.Close();

            File.WriteAllText(path, content);
        }

        public static bool ValidateTemplateFolder(string templateFolder)
        {
            if (string.IsNullOrEmpty(templateFolder))
            {
                MessageBox.Show(
                    "Please specify a template folder under Tools -> Options -> Crm DevEx -> Template Options");
                return false;
            }

            if (!Directory.Exists(templateFolder))
            {
                MessageBox.Show(
                    "Please specify a valid template folder under Tools -> Options -> Crm DevEx -> Template Options");
                return false;
            }

            return true;
        }

        public static bool ValidateTemplateFile(string templateFolder)
        {
            if (File.Exists(Path.Combine(templateFolder, ExtensionConstants.TemplateConfigFile)))
                return true;

            MessageBoxResult createResult = MessageBox.Show($"Create custom template configuration file at: {templateFolder} ? {Environment.NewLine + Environment.NewLine}" +
                                                            "Once created, add template files and try again.",
                "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
            if (createResult == MessageBoxResult.Yes)
                CreateTemplateFileTemplate(templateFolder);

            return false;
        }

        public static CustomTemplatePicker GetCustomTemplate(List<CustomTemplate> results)
        {
            CustomTemplatePicker templatePicker = new CustomTemplatePicker(results);
            bool? result = templatePicker.ShowModal();
            if (!result.HasValue || result.Value == false)
                throw new WizardBackoutException();

            if (templatePicker.SelectedTemplate == null)
                OutputLogger.WriteToOutputWindow("TemplatePicker had no selected template", MessageType.Error);

            return templatePicker;
        }
    }
}