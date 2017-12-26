using CrmDeveloperExtensions2.Core;
using EnvDTE;
using PluginDeployer.Models;
using PluginDeployer.Spkl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PluginDeployer
{
    public class SpklHelpers
    {
        public static string[] AssemblyProperties(string assemblyPath, bool isWorkflow)
        {
            AssemblyContainer assemblyContainer = null;
            try
            {
                string assemblyFolderPath = Path.GetDirectoryName(assemblyPath);

                assemblyContainer = AssemblyContainer.LoadAssembly(File.ReadAllBytes(assemblyPath), isWorkflow, assemblyFolderPath, true);

                List<PluginData> pluginDatas = assemblyContainer.PluginDatas;

                AssemblyName assemblyName = pluginDatas.First().AssemblyName;

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
                string assemblyFolderPath = Path.GetDirectoryName(assemblyPath);

                assemblyContainer = AssemblyContainer.LoadAssembly(File.ReadAllBytes(assemblyPath), isWorkflow, assemblyFolderPath, true);

                List<PluginData> pluginDatas = assemblyContainer.PluginDatas;

                return pluginDatas.First().AssemblyFullName;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        public static bool RegAttributeDefinitionExists(DTE dte, Project project)
        {
            foreach (CodeElement codeElement in project.CodeModel.CodeElements)
            {
                if (codeElement.Kind != vsCMElement.vsCMElementNamespace)
                    continue;

                // ReSharper disable once SuspiciousTypeConversion.Global
                CodeNamespace codeNamespace = codeElement as CodeNamespace;
                if (codeNamespace?.Members == null)
                    continue;

                foreach (CodeElement codeNamespaceMember in codeNamespace?.Members)
                {
                    if (codeNamespaceMember.Kind != vsCMElement.vsCMElementClass)
                        continue;

                    // ReSharper disable once SuspiciousTypeConversion.Global
                    CodeClass codeClass = codeNamespaceMember as CodeClass;
                    if (codeClass != null && codeClass.FullName.Contains(ExtensionConstants.SpklRegAttrClassName))
                        return true;
                }
            }

            return false;
        }
    }
}