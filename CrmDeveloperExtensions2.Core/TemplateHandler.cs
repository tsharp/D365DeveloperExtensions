using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using System;
using System.IO;
using System.Reflection;
using System.Windows;

namespace CrmDeveloperExtensions2.Core
{
    public class TemplateHandler
    {
        public static void AddFileFromTemplate(Project project, string templatePartialPath, string filename)
        {
            try
            {
                string codebase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(codebase, UriKind.Absolute);
                string path = Path.GetDirectoryName(uri.LocalPath);

                if (string.IsNullOrEmpty(path))
                {
                    OutputLogger.WriteToOutputWindow($"Error finding extension template directory: {path}", MessageType.Error);
                    return;
                }

                //TODO: update path for localization
                var templatePath = Path.Combine(path, $@"ItemTemplates\CSharp\Crm DevEx\1033\{templatePartialPath}.vstemplate");

                project.ProjectItems.AddFromTemplate(templatePath, filename);
            }
            catch
            {
                MessageBox.Show($"Error creating file: {filename} from template");
            }
        }
    }
}