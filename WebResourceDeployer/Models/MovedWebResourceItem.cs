using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Models
{
    public class MovedWebResourceItem
    {
        public WebResourceItem WebResourceItem { get; set; }
        public bool Publish { get; set; }
    }
}