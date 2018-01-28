using EnvDTE;

namespace D365DeveloperExtensions.Core.Connection
{
    public class ProjectItemMovedEventArgs
    {
        public string PreMoveName { get; set; }
        public ProjectItem PostMoveProjectItem { get; set; }
    }
}