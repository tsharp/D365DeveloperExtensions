using EnvDTE;

namespace CrmDeveloperExtensions2.Core.Connection
{
    public class ProjectItemMovedEventArgs
    {
        public string PreMoveName { get; set; }
        public ProjectItem PostMoveProjectItem { get; set; }
    }
}