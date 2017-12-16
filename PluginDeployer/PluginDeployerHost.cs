using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace PluginDeployer
{
    [Guid("FA0E0759-D337-4C4C-8474-217A6BDC3C06")] //Also located in ExtensionConstants.cs 
    public sealed class PluginDeployerHost : ToolWindowPane
    {
        public PluginDeployerHost() : base(null)
        {
            Caption = Resources.Resource.PluginDeployer_Window_Title;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new PluginDeployerWindow();
        }
    }
}