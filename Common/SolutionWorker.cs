using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace Common
{
    public static class SolutionWorker
    {
        public static void SetBuildConfigurationOff(SolutionConfigurations buildConfigurations, string projectName)
        {
            foreach (SolutionConfiguration buildConfiguration in buildConfigurations)
            {
                if (buildConfiguration.Name != "Debug"  && buildConfiguration.Name != "Release")
                    continue;

                SolutionContexts contexts = buildConfiguration.SolutionContexts;
                foreach (SolutionContext solutionContext in contexts)
                {
                    if (solutionContext.ProjectName == projectName)
                        solutionContext.ShouldBuild = false;
                }
            }
        }
    }
}
