using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace PluginDeployer
{
    [Guid("FA0E0759-D337-4C4C-8474-217A6BDC3C06")] //Also located in XrmToolingConnection.xaml.cs 
    public sealed class PluginDeployerHost : ToolWindowPane
    {
        public PluginDeployerHost() : base(null)
        {
            Caption = Resources.Resource.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new PluginDeployerWindow();
        }
    }
}
