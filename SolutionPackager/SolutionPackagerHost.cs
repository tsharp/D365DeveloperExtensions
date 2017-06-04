using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace SolutionPackager
{
    [Guid("F8BF1118-57B6-4404-9923-8A98AB710EBA")] //Also located in XrmToolingConnection.xaml.cs 
    public sealed class SolutionPackagerHost : ToolWindowPane
    {
        public SolutionPackagerHost() : base(null)
        {
            Caption = Resources.Resource.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new SolutionPackagerWindow();
        }
    }
}