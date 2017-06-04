using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using PluginDeployer.ViewModels;

namespace PluginDeployer
{
    public static class ModelBuilder
    {
        public static List<CrmSolution> CreateCrmSolutionView(EntityCollection solutions)
        {
            List<CrmSolution> crmSolutions = new List<CrmSolution>();

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

        private static List<CrmSolution> SortSolutions(List<CrmSolution> solutions)
        {
            //Default on top
            var i = solutions.FindIndex(s => s.SolutionId == ExtensionConstants.DefaultSolutionId);
            var item = solutions[i];
            solutions.RemoveAt(i);
            solutions.Insert(0, item);

            return solutions;
        }

        public static List<CrmAssembly> CreateCrmAssemblyView(EntityCollection assemblies)
        {
            List<CrmAssembly> crmAssemblies = new List<CrmAssembly>();

            foreach (Entity assembly in assemblies.Entities)
            {
                CrmAssembly crmAssembly = new CrmAssembly
                {
                    AssemblyId = assembly.Id,
                    Name = assembly.GetAttributeValue<string>("name"),
                    Version = Version.Parse(assembly.GetAttributeValue<string>("version")),
                    DisplayName = assembly.GetAttributeValue<string>("name") + " (" + assembly.GetAttributeValue<string>("version") + ")"
                    //,
                    //SolutionId = new Guid(assembly.GetAttributeValue<AliasedValue>("solutioncomponent.solutionid").Value.ToString())
                };

                if (assembly.Contains("plugintype.isworkflowactivity"))
                    crmAssembly.IsWorkflow = (bool)assembly.GetAttributeValue<AliasedValue>("plugintype.isworkflowactivity").Value;

                crmAssemblies.Add(crmAssembly);
            }

            return crmAssemblies;
        }
    }
}
