using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Controls
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
            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            WebBrowser.OpenUrl(dte, "https://www.paypal.me/JLattimer");
        }
    }
}