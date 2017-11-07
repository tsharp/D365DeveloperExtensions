using System;
using CrmDeveloperExtensions2.Core.Enums;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class ToolWindow
    {
        public Guid ToolWindowsId { get; set; }
        public ToolWindowType Type { get; set; }
    }
}