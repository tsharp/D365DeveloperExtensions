using EnvDTE;

namespace D365DeveloperExtensions.Core.Connection
{
    public class SolutionProjectRemovedEventArgs
    {
        public Project Project { get; set; }
    }
}