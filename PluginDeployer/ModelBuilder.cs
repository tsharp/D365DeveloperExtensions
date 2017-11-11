using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
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
            ObservableCollection<CrmSolution> crmSolutions = new ObservableCollection<CrmSolution>();

            foreach (Entity entity in solutions.Entities)
            {
                CrmSolution solution = new CrmSolution
                {
                    SolutionId = entity.Id,
                    Name = entity.GetAttributeValue<string>("friendlyname"),
                    UniqueName = entity.GetAttributeValue<string>("uniquename")
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
            ObservableCollection<CrmAssembly> crmAssemblies = new ObservableCollection<CrmAssembly>();

            foreach (Entity assembly in assemblies.Entities)
            {
                CrmAssembly crmAssembly = new CrmAssembly
                {
                    AssemblyId = assembly.Id,
                    Name = assembly.GetAttributeValue<string>("name"),
                    Version = assembly.GetAttributeValue<string>("version"),
                    DisplayName = assembly.GetAttributeValue<string>("name") + " (" + assembly.GetAttributeValue<string>("version") + ")",
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
                DisplayName = String.Empty
            });

            return crmAssemblies;
        }
    }
}