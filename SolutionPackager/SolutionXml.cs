using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using System;
using System.IO;
using System.Xml;

namespace SolutionPackager
{
    public class SolutionXml
    {
        public static bool ValidateSolutionXml(Project project)
        {
            try
            {
                if (!SolutionXmlExists(project))
                {
                    OutputLogger.WriteToOutputWindow("Solution.xml does not exist at: " + Path.GetDirectoryName(project.FullName) + "\\Other", MessageType.Error);
                    return false;
                }

                string solutionXmlPath = GetSolutionXmlPath(project);
                XmlDocument doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                XmlNodeList versionNodes = doc.GetElementsByTagName("Version");
                if (versionNodes.Count != 1)
                {
                    OutputLogger.WriteToOutputWindow("Invalid Solutions.xml: could not locate 'Versions' node", MessageType.Error);
                    return false;
                }

                Version version;
                bool validVersion = Version.TryParse(versionNodes[0].InnerText, out version);
                if (!validVersion)
                {
                    OutputLogger.WriteToOutputWindow("Invalid Solutions.xml: invalid version", MessageType.Error);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Unexpected error validating Solution.xml: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }

        public static Version GetSolutionXmlVersion(Project project)
        {
            bool isValid = ValidateSolutionXml(project);
            if (!isValid)
                return null;

            string solutionXmlPath = GetSolutionXmlPath(project);
            XmlDocument doc = new XmlDocument();
            doc.Load(solutionXmlPath);

            XmlNodeList versionNodes = doc.GetElementsByTagName("Version");

            return Version.Parse(versionNodes[0].InnerText);
        }

        public static bool SetSolutionXmlVersion(Project project, Version newVersion)
        {
            try
            {
                bool isValid = ValidateSolutionXml(project);
                if (!isValid)
                    return false;

                Version oldVersion = GetSolutionXmlVersion(project);
                if (newVersion < oldVersion)
                {
                    OutputLogger.WriteToOutputWindow("Unexpected error setting Solution.xml version: new version cannot be lower than old version", MessageType.Error);
                    return false;
                }

                string solutionXmlPath = GetSolutionXmlPath(project);
                XmlDocument doc = new XmlDocument();
                doc.Load(solutionXmlPath);

                XmlNodeList versionNodes = doc.GetElementsByTagName("Version");

                versionNodes[0].InnerText = newVersion.ToString();

                doc.Save(solutionXmlPath);

                return true;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Unexpected error setting Solution.xml version: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return false;
            }
        }

        public static bool SolutionXmlExists(Project project)
        {
            string solutionXmlPath = GetSolutionXmlPath(project);

            //TODO: check it's included in the project - not excluded but present on disk

            return File.Exists(solutionXmlPath);
        }

        private static string GetSolutionXmlPath(Project project)
        {
            string projectPath = Path.GetDirectoryName(project.FullName);
            if (String.IsNullOrEmpty(projectPath))
                return null;

            return Path.Combine(projectPath, "Other", "Solution.xml");
        }

        public static string GetLatestSolutionPath(Project project, string projectFolder)
        {
            string latestSolutionPath = null;
            string solutionProjectFolder = Packager.GetProjectSolutionFolder(project, projectFolder);

            DirectoryInfo d = new DirectoryInfo(solutionProjectFolder);
            FileInfo[] files = d.GetFiles("*.zip");

            Version currentVersion = new Version(0, 0, 0, 0);

            foreach (FileInfo file in files)
            {
                Version v = CrmDeveloperExtensions2.Core.Versioning.SolutionNameToVersion(file.Name);
                if (v <= currentVersion)
                    continue;

                latestSolutionPath = file.FullName;
                currentVersion = v;
            }

            return latestSolutionPath;
        }
    }
}