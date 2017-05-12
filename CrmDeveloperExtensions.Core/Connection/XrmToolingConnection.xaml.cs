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
using CrmDeveloperExtensions.Core;
using CrmDeveloperExtensions.Core.Connection;
using EnvDTE;

namespace CrmDeveloperExtensions.Core
{
    public partial class XrmToolingConnection : UserControl
    {
        private DTE _dte;
        public event EventHandler<Connection.ConnectEventArgs> Connected;

        public XrmToolingConnection()
        {
            InitializeComponent();
            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            GetProjectsForList();
        }

        protected virtual void OnConnected(Connection.ConnectEventArgs e)
        {
            var handler = Connected;
            handler?.Invoke(this, e);
        }

        private void GetProjectsForList()
        {
            IList<Project> projects = SolutionWorker.GetProjects();
            ProjectsDdl.ItemsSource = projects;
            ProjectsDdl.DisplayMemberPath = "Name";
        }

        private void Connect_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusBar.SetStatusBarValue(_dte, Core.Resources.Resource.StatusBarMessageConnecting, vsStatusAnimation.vsStatusAnimationGeneral);

                CrmLoginForm ctrl = new CrmLoginForm();
                ctrl.ConnectionToCrmCompleted += ctrl_ConnectionToCrmCompleted;
                ctrl.ShowDialog();

                if (ctrl.CrmConnectionMgr?.CrmSvc != null && ctrl.CrmConnectionMgr.CrmSvc.IsReady)
                {
                    MessageBox.Show("Connected to CRM! Version: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgVersion +
                                    " Org: " + ctrl.CrmConnectionMgr.CrmSvc.ConnectedOrgUniqueName, "Connection Status");
                }
                else
                {
                    MessageBox.Show("Cannot connect; try again!", "Connection Status");
                }
            }
            finally
            {
                StatusBar.ClearStatusBarValue(_dte, vsStatusAnimation.vsStatusAnimationGeneral);
            }
        }

        private void ctrl_ConnectionToCrmCompleted(object sender, EventArgs e)
        {
            if (sender is CrmLoginForm)
            {
                CrmLoginForm loginForm = (CrmLoginForm)sender;
                OnConnected(new ConnectEventArgs
                {
                    ServiceClient = loginForm.CrmConnectionMgr.CrmSvc
                });

                Dispatcher.Invoke(() =>
                {
                    ((CrmLoginForm)sender).Close();
                });
            }
        }

        private void ProjectsDdl_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }
    }
}
