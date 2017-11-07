using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class SpklConfig
    {
        public List<WebresourceDeployConfig> webresources { get; set; }
        public List<PluginDeployConfig> plugins { get; set; }
        public List<EarlyBoundTypeConfig> earlyboundtypes { get; set; }
        public List<SolutionPackageConfig> solutions { get; set; }
    }

    public class WebresourceDeployConfig
    {
        public string profile { get; set; }
        public string root { get; set; }
        public string solution { get; set; }
        public List<SpklConfigWebresourceFile> files { get; set; }
    }

    public class SpklConfigWebresourceFile
    {
        public string uniquename { get; set; }
        public string file { get; set; }
        public string description { get; set; }
    }

    public class PluginDeployConfig
    {
        public string profile { get; set; }
        public string assemblypath { get; set; }
        public string classRegex { get; set; }
    }

    public class EarlyBoundTypeConfig
    {
        public string entities { get; set; }
        public string actions { get; set; }
        public bool generateOptionsetEnums { get; set; }
        public bool generateGlobalOptionsets { get; set; }
        public bool generateStateEnums { get; set; }
        public string filename { get; set; }
        public string classNamespace { get; set; }
        public string serviceContextName { get; set; }
    }

    public class SolutionPackageConfig
    {
        public string profile { get; set; }
        public string solution_uniquename { get; set; }
        public string packagepath { get; set; }
        public string solutionpath { get; set; }
        public string packagetype { get; set; }
        public bool increment_on_import { get; set; }
        public List<SolutionPackageMap> map { get; set; }
    }

    public class SolutionPackageMap
    {
        public string map { get; set; }
        public string from { get; set; }
        public string to { get; set; }
    }
}