using System.Windows;

namespace D365DeveloperExtensions.Core.Controls
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
            WebBrowser.OpenUrl("https://github.com/scottdurow/SparkleXrm/wiki/spkl");
        }
    }
}