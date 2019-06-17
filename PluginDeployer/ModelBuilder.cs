using D365DeveloperExtensions.Core;
using Microsoft.Xrm.Sdk;
using PluginDeployer.Spkl;
using PluginDeployer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginDeployer
{
    public static class ModelBuilder
    {
        public static ObservableCollection<CrmSolution> CreateCrmSolutionView(EntityCollection solutions)
        {
            var crmSolutions = new ObservableCollection<CrmSolution>();

            foreach (var entity in solutions.Entities)
            {
                var solution = new CrmSolution
                {
                    SolutionId = entity.Id,
                    Name = entity.GetAttributeValue<string>("friendlyname"),
                    UniqueName = entity.GetAttributeValue<string>("uniquename"),
                    NameVersion = $"{entity.GetAttributeValue<string>("friendlyname")} {entity.GetAttributeValue<string>("version")}"
                };

                crmSolutions.Add(solution);
            }

            crmSolutions = SortSolutions(crmSolutions);

            return crmSolutions;
        }

        private static ObservableCollection<CrmSolution> SortSolutions(ObservableCollection<CrmSolution> solutions)
        {
            //Default on top
            var defaultSolution = solutions.FirstOrDefault(s => s.SolutionId == ExtensionConstants.DefaultSolutionId);

            solutions.Remove(defaultSolution);

            solutions.Insert(0, defaultSolution);

            return solutions;
        }

        public static ObservableCollection<CrmAssembly> CreateCrmAssemblyView(EntityCollection assemblies)
        {
            var crmAssemblies = new ObservableCollection<CrmAssembly>();

            foreach (var assembly in assemblies.Entities)
            {
                var crmAssembly = new CrmAssembly
                {
                    AssemblyId = assembly.Id,
                    Name = assembly.GetAttributeValue<string>("name"),
                    Version = assembly.GetAttributeValue<string>("version"),
                    DisplayName = $"{assembly.GetAttributeValue<string>("name")} ({assembly.GetAttributeValue<string>("version")})",
                    SolutionId = ((EntityReference)assembly.GetAttributeValue<AliasedValue>("solutioncomponent.solutionid").Value).Id
                };

                if (assembly.Contains("plugintype.isworkflowactivity"))
                    crmAssembly.IsWorkflow = (bool)assembly.GetAttributeValue<AliasedValue>("plugintype.isworkflowactivity").Value;

                crmAssemblies.Add(crmAssembly);
            }

            crmAssemblies.Insert(0, new CrmAssembly
            {
                AssemblyId = Guid.Empty,
                Name = string.Empty,
                Version = string.Empty,
                DisplayName = string.Empty
            });

            return crmAssemblies;
        }

        public static CrmAssembly CreateCrmAssembly(string projectAssemblyName, string assemblyFilePath,
            string[] assemblyProperties)
        {
            var assembly = new CrmAssembly
            {
                Name = projectAssemblyName,
                AssemblyPath = assemblyFilePath,
                Version = assemblyProperties[2],
                Culture = assemblyProperties[4],
                PublicKeyToken = assemblyProperties[6],
                //TODO: option to make none?
                IsolationMode = IsolationModeEnum.Sandbox
            };
            return assembly;
        }
    }
}