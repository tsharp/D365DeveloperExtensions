using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using EnvDTE;
using NLog;
using SolutionPackager.Resources;
using System;
using System.IO;
using System.Xml;

namespace SolutionPackager
{
    public class SolutionXml
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool ValidateSolutionXml(Project project, string projectFolder)
        {
            try
            {
                if (!SolutionXmlExists(project, projectFolder))
                {
                    OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_SolutionXmlNotExist}: {Path.GetDirectoryName(project.FullName)}\\Other", MessageType.Error);
                    return false;
                }

                var solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
                var doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                var versionNodes = doc.GetElementsByTagName("Version");
                if (versionNodes.Count != 1)
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_InvalidSolutionXml_VersionsNode, MessageType.Error);
                    return false;
                }

                var validVersion = Version.TryParse(versionNodes[0].InnerText, out _);
                if (!validVersion)
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_InvalidSolutionXml_InvalidVersion, MessageType.Error);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_UnexpectedErrorValidatingSolutionXml}: {ex.Message}" +
                    Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }

        public static Version GetSolutionXmlVersion(Project project, string projectFolder)
        {
            var isValid = ValidateSolutionXml(project, projectFolder);
            if (!isValid)
                return null;

            var solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
            var doc = new XmlDocument();
            doc.Load(solutionXmlPath);

            var versionNodes = doc.GetElementsByTagName("Version");

            return Version.Parse(versionNodes[0].InnerText);
        }

        public static bool SetSolutionXmlVersion(Project project, Version newVersion, string projectFolder)
        {
            try
            {
                var isValid = ValidateSolutionXml(project, projectFolder);
                if (!isValid)
                    return false;

                var oldVersion = GetSolutionXmlVersion(project, projectFolder);
                if (newVersion < oldVersion)
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_SolutionXmlVersionTooLow, MessageType.Error);
                    return false;
                }

                var solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
                var doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                var versionNodes = doc.GetElementsByTagName("Version");

                versionNodes[0].InnerText = newVersion.ToString();

                doc.Save(solutionXmlPath);

                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow($"{Resource.ErrorMessage_SetSolutionXmlVersion}: {ex.Message}" +
                    Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }

        public static bool SolutionXmlExists(Project project, string packageFolder)
        {
            if (string.IsNullOrEmpty(packageFolder))
                return false;

            var solutionXmlPath = GetSolutionXmlPath(project, packageFolder);

            return File.Exists(solutionXmlPath);
        }

        private static string GetSolutionXmlPath(Project project, string packageFolder)
        {
            var projectPath = Path.GetDirectoryName(project.FullName);
            if (string.IsNullOrEmpty(projectPath))
                return null;

            packageFolder = packageFolder.Replace("/", string.Empty);

            return Path.Combine(projectPath, packageFolder, "Other", "Solution.xml");
        }

        public static string GetLatestSolutionPath(Project project, string solutionFolder)
        {
            string latestSolutionPath = null;

            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            var solutionProjectFolder = Path.Combine(projectPath, solutionFolder.Replace("/", string.Empty));

            var d = new DirectoryInfo(solutionProjectFolder);
            var files = d.GetFiles("*.zip");

            var currentVersion = new Version(0, 0, 0, 0);

            foreach (var file in files)
            {
                var v = D365DeveloperExtensions.Core.Versioning.SolutionNameToVersion(file.Name);
                if (v <= currentVersion)
                    continue;

                latestSolutionPath = file.FullName;
                currentVersion = v;
            }

            return latestSolutionPath;
        }
    }
}