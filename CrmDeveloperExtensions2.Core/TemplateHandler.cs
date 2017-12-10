using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Resources;
using EnvDTE;
using NLog;
using System;
using System.IO;
using System.Reflection;
using System.Windows;

namespace CrmDeveloperExtensions2.Core
{
    public class TemplateHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void AddFileFromTemplate(Project project, string templatePartialPath, string filename)
        {
            try
            {
                string codebase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(codebase, UriKind.Absolute);
                string path = Path.GetDirectoryName(uri.LocalPath);

                if (string.IsNullOrEmpty(path))
                {
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_FindTemplateDirectory}: {path}", MessageType.Error);
                    return;
                }

                //TODO: update path for localization
                var templatePath = Path.Combine(path, $@"ItemTemplates\CSharp\Crm DevEx\1033\{templatePartialPath}.vstemplate");

                project.ProjectItems.AddFromTemplate(templatePath, filename);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, $"{Resource.ErrorMessage_CreateFileFromTemplate}: {filename}", ex);
                MessageBox.Show($"{Resource.ErrorMessage_CreateFileFromTemplate}: {filename}");
            }
        }
    }
}