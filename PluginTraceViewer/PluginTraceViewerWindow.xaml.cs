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
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            await GetCrmData();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService.CrmConnectOrgUriActual,
                $"tools/solution/edit.aspx?id=%7b{ExtensionConstants.DefaultSolutionId}%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService.CrmConnectOrgUriActual,
                "tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();
        }

        private void ResetForm()
        {
            CrmPluginTraces.ItemsSource = null;
        }

        private async Task GetCrmData()
        {
            ShowMessage("Getting plug-in trace logs...", vsStatusAnimation.vsStatusAnimationSync);

            var traceTask = GetCrmPluginTraces();

            await Task.WhenAll(traceTask);

            HideMessage(vsStatusAnimation.vsStatusAnimationSync);

            if (!traceTask.Result)
            {
                MessageBox.Show("Error Plug-in Trace Logs. See the Output Window for additional details.");
                return;
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

        private void DetailsRowClick_Handler(object sender, MouseButtonEventArgs e)
        {
            DataGridRow row = sender as DataGridRow;
            if (row != null)
                row.DetailsVisibility = row.IsSelected ? Visibility.Collapsed : Visibility.Visible;
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
    }
}