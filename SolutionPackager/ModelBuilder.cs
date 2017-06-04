using System;
using System.Collections.Generic;
using CrmDeveloperExtensions2.Core;
using Microsoft.Xrm.Sdk;
using SolutionPackager.ViewModels;

namespace SolutionPackager
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
    }
}
