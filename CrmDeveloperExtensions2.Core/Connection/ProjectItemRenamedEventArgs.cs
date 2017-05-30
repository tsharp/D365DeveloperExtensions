using EnvDTE;

namespace CrmDeveloperExtensions2.Core.Connection
{
    public class ProjectItemRenamedEventArgs
    {
        public ProjectItem ProjectItem { get; set; }
        public string OldName { get; set; }
    }
}
