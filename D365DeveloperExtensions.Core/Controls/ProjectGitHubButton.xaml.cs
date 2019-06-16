using Microsoft.VisualStudio.Shell;
using System.Windows;

namespace D365DeveloperExtensions.Core.Controls
{
    public partial class ProjectGitHubButton
    {
        public ProjectGitHubButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void OpenGitHub_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            WebBrowser.OpenUrl("https://github.com/jlattimer/D365DeveloperExtensions");
        }
    }
}