using System;
using D365DeveloperExtensions.Core.Enums;

namespace D365DeveloperExtensions.Core.Models
{
    public class ToolWindow
    {
        public Guid ToolWindowsId { get; set; }
        public ToolWindowType Type { get; set; }
    }
}