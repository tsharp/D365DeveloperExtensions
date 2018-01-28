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

                string solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
                XmlDocument doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                XmlNodeList versionNodes = doc.GetElementsByTagName("Version");
                if (versionNodes.Count != 1)
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_InvalidSolutionXml_VersionsNode, MessageType.Error);
                    return false;
                }

                bool validVersion = Version.TryParse(versionNodes[0].InnerText, out var version);
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
            bool isValid = ValidateSolutionXml(project, projectFolder);
            if (!isValid)
                return null;

            string solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
            XmlDocument doc = new XmlDocument();
            doc.Load(solutionXmlPath);

            XmlNodeList versionNodes = doc.GetElementsByTagName("Version");

            return Version.Parse(versionNodes[0].InnerText);
        }

        public static bool SetSolutionXmlVersion(Project project, Version newVersion, string projectFolder)
        {
            try
            {
                bool isValid = ValidateSolutionXml(project, projectFolder);
                if (!isValid)
                    return false;

                Version oldVersion = GetSolutionXmlVersion(project, projectFolder);
                if (newVersion < oldVersion)
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_SolutionXmlVersionTooLow, MessageType.Error);
                    return false;
                }

                string solutionXmlPath = GetSolutionXmlPath(project, projectFolder);
                XmlDocument doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                XmlNodeList versionNodes = doc.GetElementsByTagName("Version");

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

            string solutionXmlPath = GetSolutionXmlPath(project, packageFolder);

            return File.Exists(solutionXmlPath);
        }

        private static string GetSolutionXmlPath(Project project, string packageFolder)
        {
            string projectPath = Path.GetDirectoryName(project.FullName);
            if (String.IsNullOrEmpty(projectPath))
                return null;

            packageFolder = packageFolder.Replace("/", String.Empty);

            return Path.Combine(projectPath, packageFolder, "Other", "Solution.xml");
        }

        public static string GetLatestSolutionPath(Project project, string solutionFolder)
        {
            string latestSolutionPath = null;

            string projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            string solutionProjectFolder = Path.Combine(projectPath, solutionFolder.Replace("/", String.Empty));

            DirectoryInfo d = new DirectoryInfo(solutionProjectFolder);
            FileInfo[] files = d.GetFiles("*.zip");

            Version currentVersion = new Version(0, 0, 0, 0);

            foreach (FileInfo file in files)
            {
                Version v = D365DeveloperExtensions.Core.Versioning.SolutionNameToVersion(file.Name);
                if (v <= currentVersion)
                    continue;

                latestSolutionPath = file.FullName;
                currentVersion = v;
            }

            return latestSolutionPath;
        }
    }
}