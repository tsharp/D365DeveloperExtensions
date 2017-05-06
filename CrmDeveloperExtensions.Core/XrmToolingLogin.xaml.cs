using Microsoft.VisualStudio.PlatformUI;
using System;
using System.Windows;
using CrmDeveloperExtensions.Core;

namespace CrmDeveloperExtensions.Core
{
    /// <summary>
    /// Interaction logic for XrmToolingLogin.xaml
    /// </summary>
    public partial class XrmToolingLogin : DialogWindow
    {
        public XrmToolingLogin()
        {
            InitializeComponent(); 
        }

        private void Button_Click(object sender, RoutedEventArgs e)
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
