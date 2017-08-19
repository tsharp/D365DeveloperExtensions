using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using NLog;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using CrmDeveloperExtensions2.Core.Controls;
using Microsoft.VisualStudio.Imaging;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;
using Task = System.Threading.Tasks.Task;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace PluginTraceViewer
{
    public partial class PluginTraceViewerWindow : UserControl, INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private readonly BackgroundWorker _worker;
        private DateTime _lastLogDate = DateTime.MinValue;
        private ObservableCollection<CrmPluginTrace> _traces;
        public ObservableCollection<CrmPluginTrace> Traces
        {
            get => _traces;
            set
            {
                _traces = value;
                OnPropertyChanged();
            }
        }

        public PluginTraceViewerWindow()
        {
            InitializeComponent();
            DataContext = this;

            //TODO: would be better if this used a converter in xaml
            Customizations.Content = Customizations.Content.ToString().ToUpper();
            Solutions.Content = Solutions.Content.ToString().ToUpper();
            Poll.Content = Poll.Content.ToString().ToUpper();
            PollOff.Content = PollOff.Content.ToString().ToUpper();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;

            _worker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };
            _worker.DoWork += WorkerOnDoWork;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private delegate void UpdateGridDelegate(ObservableCollection<CrmPluginTrace> pluginTraces);

        private void UpdateDelegateGrid(ObservableCollection<CrmPluginTrace> newTraces)
        {
            if (newTraces.Count <= 0)
                return;

            newTraces = new ObservableCollection<CrmPluginTrace>(newTraces.OrderBy(t => t.CreatedOn));

            OutputLogger.WriteToOutputWindow("Adding new traces: " + newTraces.Count, MessageType.Info);
            foreach (CrmPluginTrace crmPluginTrace in newTraces)
                Traces.Insert(0, crmPluginTrace);

            _lastLogDate = GetLastDate();
        }

        private void WorkerOnDoWork(object sender, DoWorkEventArgs doWorkEventArgs)
        {
            while (!_worker.CancellationPending)
            {
                EntityCollection results = Crm.PluginTrace.RetrievePluginTracesFromCrm(ConnPane.CrmService, _lastLogDate);
                ObservableCollection<CrmPluginTrace> newTraces = ModelBuilder.CreateCrmPluginTraceView(results);

                if (newTraces.Count > 0)
                {
                    OutputLogger.WriteToOutputWindow("Retrieved new traces: " + newTraces.Count, MessageType.Info);

                    UpdateGridDelegate updateGridDelegate = UpdateDelegateGrid;
                    Dispatcher.BeginInvoke(DispatcherPriority.Normal, updateGridDelegate, newTraces);
                }
                else
                {
                    OutputLogger.WriteToOutputWindow("No new traces", MessageType.Info);
                }

                System.Threading.Thread.Sleep(30000);
            }
        }

        private void CancelBackgroundWorker()
        {
            _worker.CancelAsync();
            Refresh.IsEnabled = true;
            OutputLogger.WriteToOutputWindow("Stopped polling for plug-in trace log records", MessageType.Info);
        }

        private DateTime GetLastDate()
        {
            if (Traces.Count <= 0)
                return DateTime.Now;

            var crmPluginTrace = CrmPluginTraces.ItemContainerGenerator.Items[0] as CrmPluginTrace;

            return crmPluginTrace?.CreatedOn ?? DateTime.Now;
        }

        private void Poll_OnClick(object sender, RoutedEventArgs e)
        {
            if (_worker.IsBusy)
            {
                CancelBackgroundWorker();

                PollOff.Visibility = Visibility.Collapsed;
                Poll.Visibility = Visibility.Visible;
            }
            else
            {
                OutputLogger.WriteToOutputWindow("Started polling for plug-in trace log records: Interval 30 seconds", MessageType.Info);
                Refresh.IsEnabled = false;
                PollOff.Visibility = Visibility.Visible;
                Poll.Visibility = Visibility.Collapsed;

                _worker.RunWorkerAsync();
            }
        }

        private void WindowEventsOnWindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window lostFocus)
        {
            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //Should be that the tool window case closed
            if (lostFocus == null)
            {
                //_worker.CancelAsync();
                CancelBackgroundWorker();
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
            CrmPluginTraces.IsEnabled = false;
            SetButtonState(false);
            if (_worker.IsBusy)
                _worker.CancelAsync();
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
                await Task.Run(() => Crm.PluginTrace.RetrievePluginTracesFromCrm(ConnPane.CrmService, DateTime.Now.AddDays(-1)));

            if (results == null)
                return false;

            Traces = ModelBuilder.CreateCrmPluginTraceView(results);
            CrmPluginTraces.IsEnabled = true;
            Refresh.Opacity = 1;

            _lastLogDate = GetLastDate();

            return true;
        }

        private void ShowMessage(string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.SetStatusBarValue(_dte, "Retrieving plug-in trace logs...", (vsStatusAnimation)animation);

                        Overlay.Show(message);
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
                        Overlay.Hide();
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
            Refresh.IsEnabled = enabled;
            Poll.IsEnabled = enabled;
            PollOff.IsEnabled = enabled;
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

        private void Delete_OnClick(object sender, RoutedEventArgs e)
        {
            Guid[] pluginTraceLogsToDelete =
                Traces.Where(d => d.PendingDelete).Select(d => d.PluginTraceLogidId).ToArray();

            List<Guid> deletedPluginTraceLogIds =
                Crm.PluginTrace.DeletePluginTracesFromCrm(ConnPane.CrmService, pluginTraceLogsToDelete);

            foreach (Guid pluginTraceLogId in deletedPluginTraceLogIds)
            {
                var pluginTraceLog = Traces.FirstOrDefault(t => t.PluginTraceLogidId == pluginTraceLogId);
                Traces.Remove(pluginTraceLog);
            }
        }
    }
}