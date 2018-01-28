using EnvDTE;

namespace D365DeveloperExtensions.Core.Connection
{
    public class SolutionProjectRenamedEventArgs
    {
        public Project Project { get; set; }
        public string OldName { get; set; }
    }
}