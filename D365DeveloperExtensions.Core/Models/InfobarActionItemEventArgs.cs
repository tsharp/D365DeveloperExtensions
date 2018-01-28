using Microsoft.VisualStudio.Shell.Interop;

namespace D365DeveloperExtensions.Core.Models
{
    public class InfobarActionItemEventArgs
    {
        public IVsInfoBarUIElement InfoBarElement { get; set; }
        public IVsInfoBarActionItem InfobarActionItem { get; set; }
    }
}
