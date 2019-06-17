using D365DeveloperExtensions.Core;
using EnvDTE;
using PluginDeployer.Spkl;
using System;
using System.IO;
using System.Linq;

namespace PluginDeployer
{
    public class SpklHelpers
    {
        public static string[] AssemblyProperties(string assemblyPath, bool isWorkflow)
        {
            AssemblyContainer assemblyContainer = null;
            try
            {
                var assemblyFolderPath = Path.GetDirectoryName(assemblyPath);

                assemblyContainer = AssemblyContainer.LoadAssembly(File.ReadAllBytes(assemblyPath), isWorkflow, assemblyFolderPath);

                var pluginDatas = assemblyContainer.PluginDatas;

                var assemblyName = pluginDatas.First().AssemblyName;

                var assemblyProperties = assemblyName.FullName.Split(",= ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                return assemblyProperties;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        public static string AssemblyFullName(string assemblyPath, bool isWorkflow)
        {
            AssemblyContainer assemblyContainer = null;
            try
            {
                var assemblyFolderPath = Path.GetDirectoryName(assemblyPath);

                assemblyContainer = AssemblyContainer.LoadAssembly(File.ReadAllBytes(assemblyPath), isWorkflow, assemblyFolderPath);

                var pluginDatas = assemblyContainer.PluginDatas;

                return pluginDatas.First().AssemblyFullName;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        public static bool RegAttributeDefinitionExists(Project project)
        {
            foreach (CodeElement codeElement in project.CodeModel.CodeElements)
            {
                if (codeElement.Kind != vsCMElement.vsCMElementNamespace)
                    continue;

                // ReSharper disable once SuspiciousTypeConversion.Global
                var codeNamespace = codeElement as CodeNamespace;
                if (codeNamespace?.Members == null)
                    continue;

                foreach (CodeElement codeNamespaceMember in codeNamespace?.Members)
                {
                    if (codeNamespaceMember.Kind != vsCMElement.vsCMElementClass)
                        continue;

                    // ReSharper disable once SuspiciousTypeConversion.Global
                    if (codeNamespaceMember is CodeClass codeClass &&
                        codeClass.FullName.Contains(ExtensionConstants.SpklRegAttrClassName))
                        return true;
                }
            }

            return false;
        }
    }
}