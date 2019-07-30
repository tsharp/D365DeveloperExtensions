using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;
using NLog;
using System;
using System.IO;
using System.Reflection;
using System.Windows;

namespace D365DeveloperExtensions.Core
{
    public class TemplateHandler
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void AddFileFromTemplate(Project project, string templatePartialPath, string filename)
        {
            try
            {
                var codebase = Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(codebase, UriKind.Absolute);
                var path = Path.GetDirectoryName(uri.LocalPath);

                if (string.IsNullOrEmpty(path))
                {
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_FindTemplateDirectory}: {path}", MessageType.Error);
                    return;
                }

                //TODO: update path for localization
                var templatePath = Path.Combine(path, $@"ItemTemplates\CSharp\D365 DevEx\1033\{templatePartialPath}.vstemplate");

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