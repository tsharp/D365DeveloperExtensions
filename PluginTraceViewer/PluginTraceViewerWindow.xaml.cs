using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Config;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Vs;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using Microsoft.VisualStudio;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using PluginTraceViewer.ViewModels;
using static CrmDeveloperExtensions2.Core.FileSystem;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;
using Task = System.Threading.Tasks.Task;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace PluginTraceViewer
{
    public partial class PluginTraceViewerWindow : UserControl
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();


        public PluginTraceViewerWindow()
        {
            InitializeComponent();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
        }

        private void WindowEventsOnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //WindowEventsOnWindowActivated in this project can be called when activating another window
            //so we don't want to contine further unless our window is active
            if (!HostWindow.IsCrmDevExWindow(gotFocus))
                return;

            //Grid is populated already
            if (CrmPluginTraces.ItemsSource != null)
                return;

            if (ConnPane.CrmService != null && ConnPane.CrmService.IsReady)
            {
                SetWindowCaption(gotFocus.Caption);
                SetButtonState(true);
                LoadData();
            }
        }

        private void SetWindowCaption(string currentCaption)
        {
            _dte.ActiveWindow.Caption = HostWindow.SetCaption(currentCaption, ConnPane.CrmService);
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            SetWindowCaption(_dte.ActiveWindow.Caption);
            SetButtonState(true);
            LoadData();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService,
                $"tools/solution/edit.aspx?id=%7b{ExtensionConstants.DefaultSolutionId}%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService,
                "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();
        }

        private void ResetForm()
        {
            CrmPluginTraces.ItemsSource = null;
            SetButtonState(false);
        }

        private async Task GetCrmData()
        {
            try
            {
                ShowMessage("Getting plug-in trace logs...", vsStatusAnimation.vsStatusAnimationSync);

                var traceTask = GetCrmPluginTraces();

                await Task.WhenAll(traceTask);

                if (!traceTask.Result)
                {
                    HideMessage(vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Plug-in Trace Logs. See the Output Window for additional details.");
                }
            }
            finally
            {
                HideMessage(vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task<bool> GetCrmPluginTraces()
        {
            EntityCollection results =
                await Task.Run(() => Crm.PluginTrace.RetrievePluginTracesFromCrm(ConnPane.CrmService));
            if (results == null)
                return false;

            List<CrmPluginTrace> pluginTraces = ModelBuilder.CreateCrmPluginTraceView(results);

            CrmPluginTraces.ItemsSource = pluginTraces;

            return true;
        }

        private void ShowMessage(string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.SetStatusBarValue(_dte, "Retrieving plug-in trace logs...", (vsStatusAnimation)animation);
                        LockMessage.Content = message;
                        LockOverlay.Visibility = Visibility.Visible;
                    }
                ));
        }

        private void HideMessage(vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.ClearStatusBarValue(_dte, (vsStatusAnimation)animation);
                        LockOverlay.Visibility = Visibility.Hidden;
                    }
                ));
        }

        private void ViewDetails_OnClick(object sender, RoutedEventArgs e)
        {
            for (var vis = sender as Visual; vis != null; vis = VisualTreeHelper.GetParent(vis) as Visual)
            {
                if (!(vis is DataGridRow))
                    continue;

                var row = (DataGridRow)vis;
                row.DetailsVisibility = row.DetailsVisibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
                break;
            }
        }

        private void SetButtonState(bool enabled)
        {
            Customizations.IsEnabled = enabled;
            Solutions.IsEnabled = enabled;
        }

        private void Refresh_OnClick(object sender, RoutedEventArgs e)
        {
            LoadData();
        }

        private async void LoadData()
        {
            await GetCrmData();
        }

        private void OpenInCrm_OnClick(object sender, RoutedEventArgs e)
        {
            Guid pluginTraceLogId = new Guid(((Button)sender).CommandParameter.ToString());

            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService, $"userdefined/edit.aspx?etc=4619&id=%7b{pluginTraceLogId}%7d");
        }
    }
}