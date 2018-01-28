using EnvDTE;

namespace D365DeveloperExtensions.Core.Models
{
    public class MovedProjectItem
    {
        public ProjectItem ProjectItem { get; set; }
        public string OldName { get; set; }
        public uint ProjectItemId { get; set; }
    }
}