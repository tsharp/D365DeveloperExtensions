using EnvDTE;

namespace D365DeveloperExtensions.Core.Connection
{
    public class ProjectItemRenamedEventArgs
    {
        public ProjectItem ProjectItem { get; set; }
        public string OldName { get; set; }
    }
}