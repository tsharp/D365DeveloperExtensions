using EnvDTE;

namespace D365DeveloperExtensions.Core.Connection
{
    public class ProjectItemRemovedEventArgs
    {
        public ProjectItem ProjectItem { get; set; }
    }
}