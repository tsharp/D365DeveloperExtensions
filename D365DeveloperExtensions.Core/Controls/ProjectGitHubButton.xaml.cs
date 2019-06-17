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
            WebBrowser.OpenUrl("https://github.com/jlattimer/D365DeveloperExtensions");
        }
    }
}