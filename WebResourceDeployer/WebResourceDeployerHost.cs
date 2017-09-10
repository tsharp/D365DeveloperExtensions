using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace WebResourceDeployer
{
    [Guid("A3479AE0-5F4F-4A14-96F4-46F39000023A")] //Also located in ExtensionConstants.cs 
    public sealed class WebResourceDeployerHost : ToolWindowPane
    {
        public WebResourceDeployerHost() : base(null)
        {
            Caption = Resources.Resource.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new WebResourceDeployerWindow();
        }
    }
}