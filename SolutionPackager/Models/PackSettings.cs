using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using SolutionPackager.ViewModels;
using System;

namespace SolutionPackager.Models
{
    public class PackSettings
    {
        public Project Project { get; set; }
        public CrmSolution CrmSolution { get; set; }
        public SolutionPackageConfig SolutionPackageConfig { get; set; }
        public string PackageFolder { get; set; }
        public bool EnablePackagerLogging { get; set; }
        public bool SaveSolutions { get; set; }
        public string SolutionFolder { get; set; }
        public string ProjectPath { get; set; }
        public string ProjectPackageFolder { get; set; }
        public string ProjectSolutionFolder { get; set; }
        public string FileName { get; set; }
        public string FullFilePath { get; set; }
        public Version Version { get; set; }
        public bool UseMapFile { get; set; }
    }
}