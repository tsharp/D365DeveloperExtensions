using D365DeveloperExtensions.Core.Models;
using EnvDTE;
using SolutionPackager.ViewModels;
using System.IO;

namespace SolutionPackager.Models
{
    public class UnpackSettings
    {
        public Project Project { get; set; }
        public CrmSolution CrmSolution { get; set; }
        public SolutionPackageConfig SolutionPackageConfig { get; set; }
        public bool EnablePackagerLogging { get; set; }
        public bool SaveSolutions { get; set; }
        public string SolutionFolder { get; set; }
        public string ProjectPath { get; set; }
        public string PackageFolder { get; set; }
        public string ProjectPackageFolder { get; set; }
        public string ProjectSolutionFolder { get; set; }
        public string DownloadedZipPath { get; set; }
        public DirectoryInfo ExtractedFolder { get; set; }
        public bool UseMapFile { get; set; }
    }
}