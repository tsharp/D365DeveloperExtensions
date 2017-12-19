using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using SolutionPackager.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace SolutionPackager
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
                    UniqueName = entity.GetAttributeValue<string>("uniquename"),
                    Version = Version.Parse(entity.GetAttributeValue<string>("version")),
                    Prefix = entity.GetAttributeValue<AliasedValue>("publisher.customizationprefix").Value.ToString(),
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
    }
}