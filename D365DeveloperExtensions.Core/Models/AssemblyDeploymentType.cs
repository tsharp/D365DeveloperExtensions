using System.Collections.Generic;

namespace D365DeveloperExtensions.Core.Models
{
    public class AssemblyDeploymentType
    {
        public string Name { get; set; }
        public int Value { get; set; }
    }

    public static class AssemblyDeploymentTypes
    {
        public static List<AssemblyDeploymentType> Types => new List<AssemblyDeploymentType> {
            new AssemblyDeploymentType{Name = "Assembly Only", Value = 0},
            new AssemblyDeploymentType{Name = "Spkl", Value = 1}
        };
    }
}