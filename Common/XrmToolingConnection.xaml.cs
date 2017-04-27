using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.VisualStudio.Shell;
using WpfApplication2;

namespace Common
{
    /// <summary>
    /// Interaction logic for XrmToolingConnection.xaml
    /// </summary>
    public partial class XrmToolingConnection : UserControl
    {
        public XrmToolingConnection()
        {
            InitializeComponent();
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            // Establish the Login control.
            CRMLoginForm1 ctrl = new CRMLoginForm1();

            // Wire event to login response. 
            ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;

            // Show the login control. 
            ctrl.ShowDialog();

            // Handle the returned CRM connection object.
            // On successful connection, display the CRM version and connected org name 
            if (ctrl.CrmConnectionMgr?.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
            {
                MessageBox.Show("Connected to CRM! Version: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion +
                                " Org: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, "Connection Status");

                // Perform your actions here
            }
            else
            {
                MessageBox.Show("Cannot connect; try again!", "Connection Status");
            }
        }

        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CRMLoginForm1)
            {
                Dispatcher.Invoke(() =>
                {
                    ((CRMLoginForm1)sender).Close();
                });
            }
        }
    }
}
