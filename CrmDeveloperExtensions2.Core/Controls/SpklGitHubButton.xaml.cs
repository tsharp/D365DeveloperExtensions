using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class SpklGitHubButton
    {
        public SpklGitHubButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void OpenGitHub_Click(object sender, RoutedEventArgs e)
        {
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            WebBrowser.OpenUrl(dte, "https://github.com/scottdurow/SparkleXrm/wiki/spkl");
        }
    }
}