using Microsoft.VisualStudio.Shell;
using System.Windows;

namespace D365DeveloperExtensions.Core.Controls
{
    public partial class PayPalButton
    {
        public PayPalButton()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void OpenPayPal_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            WebBrowser.OpenUrl("https://www.paypal.me/JLattimer");
        }
    }
}