using EnvDTE;

namespace CrmDeveloperExtensions2.Core.Connection
{
    public class SolutionProjectRenamedEventArgs
    {
        public Project Project { get; set; }
        public string OldName { get; set; }
    }
}
