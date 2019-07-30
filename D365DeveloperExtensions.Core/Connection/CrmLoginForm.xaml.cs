using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Tooling.CrmConnectControl;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;

namespace D365DeveloperExtensions.Core.Connection
{
    public partial class CrmLoginForm
    {
        #region Vars
        /// <summary>
        /// Microsoft.Xrm.Tooling.Connector services
        /// </summary>
        private CrmServiceClient _crmSvc;
        /// <summary>
        /// Bool flag to determine if there is a connection 
        /// </summary>
        private bool _bIsConnectedComplete;
        /// <summary>
        /// CRM Connection Manager component. 
        /// </summary>
        private CrmConnectionManager _mgr;
        /// <summary>
        ///  This is used to allow the UI to reset w/out closing 
        /// </summary>
        private bool _resetUiFlag;

        private readonly bool _autoLogin;
        #endregion

        #region Properties
        /// <summary>
        /// CRM Connection Manager 
        /// </summary>
        public CrmConnectionManager CrmConnectionMgr => _mgr;

        #endregion

        #region Event
        /// <summary>
        /// Raised when a connection to CRM has completed. 
        /// </summary>
        public event EventHandler ConnectionToCrmCompleted;
        #endregion


        public CrmLoginForm(bool autoLogin)
        {
            InitializeComponent();
            //// Should be used for testing only.
            //ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) =>
            //{
            //    MessageBox.Show("CertError");
            //    return true;
            //};
            _autoLogin = autoLogin;
            EnableXrmToolingLogging();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            /*
				This is the setup process for the login control, 
				The login control uses a class called CrmConnectionManager to manage the interaction with CRM, this class and also be queried as later points for information about the current connection. 
				In this case, the login control is referred to as CrmLoginCtrl
			 */

            // Set off flag. 
            _bIsConnectedComplete = false;

            // Init the CRM Connection manager.. 
            _mgr = new CrmConnectionManager();
            // Pass a reference to the current UI or container control,  this is used to synchronize UI threads In the login control
            _mgr.ParentControl = CrmLoginCtrl;
            // if you are using an unmanaged client, excel for example, and need to store the config in the users local directory
            // set this option to true. 
            _mgr.UseUserLocalDirectoryForConfigStore = true;

            _mgr.ClientId = "2ad88395-b77d-4561-9441-d0e40824f9bc";
            _mgr.RedirectUri = new Uri("app://5d3e90d6-aa8e-48a8-8f2c-58b45cc67315");
            // if you are using an unmanaged client,  you need to provide the name of an exe to use to create app config key's for. 
            //mgr.HostApplicatioNameOveride = "MyExecName.exe";
            // CrmLoginCtrl is the Login control,  this sets the CrmConnection Manager into it. 
            CrmLoginCtrl.SetGlobalStoreAccess(_mgr);
            // There are several modes to the login control UI
            CrmLoginCtrl.SetControlMode(ServerLoginConfigCtrlMode.FullLoginPanel);
            // this wires an event that is raised when the login button is pressed. 
            CrmLoginCtrl.ConnectionCheckBegining += CrmLoginCtrl_ConnectionCheckBegining;
            // this wires an event that is raised when an error in the connect process occurs. 
            CrmLoginCtrl.ConnectErrorEvent += CrmLoginCtrl_ConnectErrorEvent;
            // this wires an event that is raised when a status event is returned. 
            CrmLoginCtrl.ConnectionStatusEvent += CrmLoginCtrl_ConnectionStatusEvent;
            // this wires an event that is raised when the user clicks the cancel button. 
            CrmLoginCtrl.UserCancelClicked += CrmLoginCtrl_UserCancelClicked;
            // Check to see if its possible to do an Auto Login 
            if (_mgr.RequireUserLogin())
                return;

            MessageBoxResult result = MessageBoxResult.No;
            if (!_autoLogin)
            {
                result = MessageBox.Show(
                      "Credentials already saved in configuration\nChoose Yes to Auto Login or No to Reset Credentials",
                      "Auto Login", MessageBoxButton.YesNo, MessageBoxImage.Question);
            }

            if (_autoLogin || result == MessageBoxResult.Yes)
                DoLogin();
        }

        private void DoLogin()
        {
            // If RequireUserLogin is false, it means that there has been a successful login here before and the credentials are cached. 
            CrmLoginCtrl.IsEnabled = false;
            // When running an auto login,  you need to wire and listen to the events from the connection manager.
            // Run Auto User Login process, Wire events. 
            _mgr.ServerConnectionStatusUpdate += mgr_ServerConnectionStatusUpdate;
            _mgr.ConnectionCheckComplete += mgr_ConnectionCheckComplete;
            // Start the connection process. 
            _mgr.ConnectToServerCheck();

            // Show the message grid. 
            CrmLoginCtrl.ShowMessageGrid();
        }

        #region Events

        private void mgr_ServerConnectionStatusUpdate(object sender, ServerConnectStatusEventArgs e)
        {
            // The Status event will contain information about the current login process,  if Connected is false, then there is not yet a connection. 
            // Set the updated status of the loading process. 
            Dispatcher.Invoke(DispatcherPriority.Normal,
                               new Action(() =>
                               {
                                   Title = string.IsNullOrWhiteSpace(e.StatusMessage) ? e.ErrorMessage : e.StatusMessage;
                               }));
        }

        private void mgr_ConnectionCheckComplete(object sender, ServerConnectStatusEventArgs e)
        {
            // The Status event will contain information about the current login process,  if Connected is false, then there is not yet a connection. 
            // Unwire events that we are not using anymore, this prevents issues if the user uses the control after a failed login. 
            ((CrmConnectionManager)sender).ConnectionCheckComplete -= mgr_ConnectionCheckComplete;
            ((CrmConnectionManager)sender).ServerConnectionStatusUpdate -= mgr_ServerConnectionStatusUpdate;

            if (!e.Connected)
            {
                // if its not connected pop the login screen here. 
                if (e.MultiOrgsFound)
                    MessageBox.Show("Unable to Login to CRM using cached credentials. Org Not found", "Login Failure");
                else
                    MessageBox.Show("Unable to Login to CRM using cached credentials", "Login Failure");

                _resetUiFlag = true;
                CrmLoginCtrl.GoBackToLogin();
                // Bad Login Get back on the UI. 
                Dispatcher.Invoke(DispatcherPriority.Normal,
                       new Action(() =>
                       {
                           Title = "Failed to Login with cached credentials.";
                           MessageBox.Show(Title, "Notification from ConnectionManager", MessageBoxButton.OK, MessageBoxImage.Error);
                           CrmLoginCtrl.IsEnabled = true;
                       }));

                _resetUiFlag = false;
            }
            else
            {
                // Good Login Get back on the UI 
                if (e.Connected && !_bIsConnectedComplete)
                    ProcessSuccess();

                OutputLogger.WriteToOutputWindow(HostWindow.GetCaption(string.Empty, _mgr.CrmSvc).Substring(3), MessageType.Info);
            }
        }

        private void CrmLoginCtrl_ConnectionCheckBegining(object sender, EventArgs e)
        {
            _bIsConnectedComplete = false;
            Dispatcher.Invoke(DispatcherPriority.Normal,
                               new Action(() =>
                               {
                                   Title = "Starting Login Process. ";
                                   CrmLoginCtrl.IsEnabled = true;
                               }));
        }

        private void CrmLoginCtrl_ConnectionStatusEvent(object sender, ConnectStatusEventArgs e)
        {
            //Here we are using the bIsConnectedComplete bool to check to make sure we only process this call once. 
            if (e.ConnectSucceeded && !_bIsConnectedComplete)
                ProcessSuccess();
        }

        private void CrmLoginCtrl_ConnectErrorEvent(object sender, ConnectErrorEventArgs e)
        {
            //MessageBox.Show(e.ErrorMessage, "Error here");
        }

        private void CrmLoginCtrl_UserCancelClicked(object sender, EventArgs e)
        {
            if (!_resetUiFlag)
                Close();
        }

        #endregion

        private void ProcessSuccess()
        {
            _resetUiFlag = true;
            _bIsConnectedComplete = true;
            _crmSvc = _mgr.CrmSvc;
            CrmLoginCtrl.GoBackToLogin();
            Dispatcher.Invoke(DispatcherPriority.Normal,
               new Action(() =>
                   {
                       Title = "Notification from Parent";
                       CrmLoginCtrl.IsEnabled = true;
                   }));

            SetGlobalConnection(_mgr.CrmSvc);

            // Notify Caller that we are done with success. 
            ConnectionToCrmCompleted?.Invoke(this, null);

            _resetUiFlag = false;
        }

        private static void EnableXrmToolingLogging()
        {
            if (!UserOptionsHelper.GetOption<bool>(UserOptionProperties.XrmToolingLoggingEnabled))
                return;

            TraceControlSettings.TraceLevel = SourceLevels.All;
            var logPath = XrmToolingLogging.GetLogFilePath();
            TraceControlSettings.AddTraceListener(new TextWriterTraceListener(logPath));
        }

        private static void SetGlobalConnection(CrmServiceClient client)
        {
            SharedGlobals.SetGlobal("CrmService", client);
        }
    }

    #region system.diagnostics settings for this control

    // Add or merge this section to your app to enable diagnostics on the use of the CRM login control and connection
    /*
  <system.diagnostics>
    <trace autoflush="true" />
    <sources>
      <source name="Microsoft.Xrm.Tooling.Connector.CrmServiceClient"
              switchName="Microsoft.Xrm.Tooling.Connector.CrmServiceClient"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>
      <source name="Microsoft.Xrm.Tooling.CrmConnectControl"
              switchName="Microsoft.Xrm.Tooling.CrmConnectControl"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>

      <source name="Microsoft.Xrm.Tooling.WebResourceUtility"
              switchName="Microsoft.Xrm.Tooling.WebResourceUtility"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>
      
    <!-- WCF DEBUG SOURCES -->
      <source name="System.IdentityModel" switchName="System.IdentityModel">
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- Log all messages in the 'Messages' tab of SvcTraceViewer. -->
      <source name="System.ServiceModel.MessageLogging" switchName="System.ServiceModel.MessageLogging" >
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- ActivityTracing and propogateActivity are used to flesh out the 'Activities' tab in
           SvcTraceViewer to aid debugging. -->
      <source name="System.ServiceModel" switchName="System.ServiceModel" propagateActivity="true">
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- END WCF DEBUG SOURCES -->
    </sources>
    <switches>
      <!-- 
            Possible values for switches: Off, Error, Warining, Info, Verbose
                Verbose:    includes Error, Warning, Info, Trace levels
                Info:       includes Error, Warning, Info levels
                Warning:    includes Error, Warning levels
                Error:      includes Error level
        -->
      <add name="Microsoft.Xrm.Tooling.Connector.CrmServiceClient" value="Verbose" />
      <add name="Microsoft.Xrm.Tooling.CrmConnectControl" value="Verbose"/>
      <add name="Microsoft.Xrm.Tooling.WebResourceUtility" value="Verbose" />
      <add name="System.IdentityModel" value="Verbose"/>
      <add name="System.ServiceModel.MessageLogging" value="Verbose"/>
      <add name="System.ServiceModel" value="Error, ActivityTracing"/>
      
    </switches>
    <sharedListeners>
      <add name="fileListener" type="System.Diagnostics.TextWriterTraceListener" initializeData="LoginControlTesterLog.txt"/>
      <!--<add name="eventLogListener" type="System.Diagnostics.EventLogTraceListener" initializeData="CRM UII"/>-->
      <add name="xml" type="System.Diagnostics.XmlWriterTraceListener" initializeData="CrmToolBox.svclog" />
    </sharedListeners>
  </system.diagnostics>
*/

    #endregion
}
