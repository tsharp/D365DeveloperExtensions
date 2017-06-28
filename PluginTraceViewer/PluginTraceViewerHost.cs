using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace PluginTraceViewer
{
    [Guid("E7A15FDA-6C33-48F8-A1E7-D78E49458A7A")] //Also located in ExtensionConstants.cs 
    public sealed class PluginTraceViewerHost : ToolWindowPane
    {
        public PluginTraceViewerHost() : base(null)
        {
            Caption = Resources.Resource.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new PluginTraceViewerWindow();
        }
    }
}