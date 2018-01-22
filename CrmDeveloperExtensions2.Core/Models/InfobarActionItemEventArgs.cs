using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class InfobarActionItemEventArgs
    {
        public IVsInfoBarUIElement InfoBarElement { get; set; }
        public IVsInfoBarActionItem InfobarActionItem { get; set; }
    }
}
