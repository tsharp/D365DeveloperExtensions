using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Connection;
using CrmDeveloperExtensions2.Core.DataGrid;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using NLog;
using PluginTraceViewer.Models;
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
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Task = System.Threading.Tasks.Task;
using WebBrowser = CrmDeveloperExtensions2.Core.WebBrowser;

namespace PluginTraceViewer
{
    public partial class PluginTraceViewerWindow : INotifyPropertyChanged
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private static readonly Logger ExtensionLogger = LogManager.GetCurrentClassLogger();
        private readonly BackgroundWorker _worker;
        private DateTime _lastLogDate = DateTime.MinValue;
        private ObservableCollection<CrmPluginTrace> _traces;
        private ObservableCollection<FilterEntity> _filterEntities;
        private ObservableCollection<FilterMessage> _filterMessages;
        private ObservableCollection<FilterMode> _filterModes;
        private ObservableCollection<FilterTypeName> _filterTypeNames;

        public ObservableCollection<CrmPluginTrace> Traces
        {
            get => _traces;
            set
            {
                _traces = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterEntity> FilterEntities
        {
            get => _filterEntities;
            set
            {
                _filterEntities = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterMessage> FilterMessages
        {
            get => _filterMessages;
            set
            {
                _filterMessages = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterMode> FilterModes
        {
            get => _filterModes;
            set
            {
                _filterModes = value;
                OnPropertyChanged();
            }
        }
        public ObservableCollection<FilterTypeName> FilterTypeNames
        {
            get => _filterTypeNames;
            set
            {
                _filterTypeNames = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public PluginTraceViewerWindow()
        {
            InitializeComponent();
            DataContext = this;

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
                EntityCollection results =
                    Crm.PluginTrace.RetrievePluginTracesFromCrm(ConnPane.CrmService, _lastLogDate);
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

            CrmPluginTrace crmPluginTrace = Traces.FirstOrDefault();

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
                OutputLogger.WriteToOutputWindow("Started polling for plug-in trace log records: Interval 30 seconds",
                    MessageType.Info);
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
            FilterEntities = new ObservableCollection<FilterEntity>();
            FilterMessages = new ObservableCollection<FilterMessage>();
            FilterModes = new ObservableCollection<FilterMode>();
            FilterTypeNames = new ObservableCollection<FilterTypeName>();

            SetWindowCaption(_dte.ActiveWindow.Caption);
            SetButtonState(true);
            LoadData();
        }

        private void ConnPane_OnSolutionBeforeClosing(object sender, EventArgs e)
        {
            ResetForm();

            ClearConnection();
        }

        private void ConnPane_OnSolutionOpened(object sender, EventArgs e)
        {
            ClearConnection();
        }

        private void ClearConnection()
        {
            ConnPane.IsConnected = false;
            ConnPane.CrmService?.Dispose();
            ConnPane.CrmService = null;
        }

        private void ResetForm()
        {
            FilterEntities = new ObservableCollection<FilterEntity>();
            FilterMessages = new ObservableCollection<FilterMessage>();
            FilterModes = new ObservableCollection<FilterMode>();
            FilterTypeNames = new ObservableCollection<FilterTypeName>();

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
                Overlay.ShowMessage(_dte, "Getting plug-in trace logs...", vsStatusAnimation.vsStatusAnimationSync);

                var traceTask = GetCrmPluginTraces();

                await Task.WhenAll(traceTask);

                if (!traceTask.Result)
                {
                    Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
                    MessageBox.Show("Error Plug-in Trace Logs. See the Output Window for additional details.");
                }
            }
            finally
            {
                Overlay.HideMessage(_dte, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private async Task<bool> GetCrmPluginTraces()
        {
            EntityCollection results =
                await Task.Run(() =>
                    Crm.PluginTrace.RetrievePluginTracesFromCrm(ConnPane.CrmService, DateTime.Now.AddDays(-1)));

            if (results == null)
                return false;

            Traces = ModelBuilder.CreateCrmPluginTraceView(results);
            CrmPluginTraces.IsEnabled = true;
            Refresh.Opacity = 1;

            CreateFilterEntityList();
            CreateFilterMessageList();
            CreateFilterModeList();
            CreateFilterTypeNameList();

            _lastLogDate = GetLastDate();

            return true;
        }

        private void CreateFilterEntityList()
        {
            FilterEntities = FilterEntity.CreateFilterList(Traces);

            foreach (FilterEntity filterEntity in FilterEntities)
            {
                filterEntity.PropertyChanged += FilterEntity_PropertyChanged;
            }

            FilterByEntity.IsEnabled = FilterEntities.Count > 2;
        }

        private void CreateFilterMessageList()
        {
            FilterMessages = FilterMessage.CreateFilterList(Traces);

            foreach (FilterMessage filterMessage in FilterMessages)
            {
                filterMessage.PropertyChanged += FilterMessage_PropertyChanged;
            }

            FilterByMessage.IsEnabled = FilterMessages.Count > 2;
        }

        private void CreateFilterModeList()
        {
            FilterModes = FilterMode.CreateFilterList(Traces);

            foreach (FilterMode filterMode in FilterModes)
            {
                filterMode.PropertyChanged += FilterMode_PropertyChanged;
            }

            FilterByMode.IsEnabled = FilterModes.Count > 2;
        }

        private void CreateFilterTypeNameList()
        {
            FilterTypeNames = FilterTypeName.CreateFilterList(Traces);

            foreach (FilterTypeName filterTypeName in FilterTypeNames)
            {
                filterTypeName.PropertyChanged += FilterTypeName_PropertyChanged;
            }

            FilterByTypeName.IsEnabled = FilterTypeNames.Count > 2;
        }

        private void ViewDetails_OnClick(object sender, RoutedEventArgs e)
        {
            DetailsRow.ShowHideDetailsRow(sender);
        }

        private void SetButtonState(bool enabled)
        {
            Refresh.IsEnabled = enabled;
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

            WebBrowser.OpenCrmPage(_dte, ConnPane.CrmService,
                $"userdefined/edit.aspx?etc=4619&id=%7b{pluginTraceLogId}%7d");
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

        private void CrmPluginTraces_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Make rows unselectable
            CrmPluginTraces.UnselectAllCells();
        }

        public void FilterTraces()
        {
            ICollectionView icv = CollectionViewSource.GetDefaultView(CrmPluginTraces.ItemsSource);
            if (icv == null) return;

            icv.Filter = GetFilteredView;
        }

        public bool GetFilteredView(object sourceObject)
        {
            CrmPluginTrace crmPluginTrace = (CrmPluginTrace)sourceObject;

            if (EntityCondition(crmPluginTrace) && MessageCondition(crmPluginTrace) && ModeCondition(crmPluginTrace) &&
                (CorrelationIdCondition(crmPluginTrace) || DetailsCondition(crmPluginTrace)) && TypeNameCondition(crmPluginTrace))
            {
                return true;
            }
            return false;
        }

        public bool EntityCondition(CrmPluginTrace crmPluginTrace)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterEntities), crmPluginTrace.Entity);
        }

        public bool MessageCondition(CrmPluginTrace crmPluginTrace)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterMessages), crmPluginTrace.MessageName);
        }

        public bool ModeCondition(CrmPluginTrace crmPluginTrace)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterModes), crmPluginTrace.Mode);
        }

        public bool TypeNameCondition(CrmPluginTrace crmPluginTrace)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(FilterTypeNames), crmPluginTrace.TypeName);
        }

        public bool DetailsCondition(CrmPluginTrace crmPluginTrace)
        {
            if (string.IsNullOrEmpty(crmPluginTrace.Details))
                return false;

            string search = DetailsSearch.Text.Trim();
            if (string.IsNullOrEmpty(search))
                return true;

            return crmPluginTrace.Details.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public bool CorrelationIdCondition(CrmPluginTrace crmPluginTrace)
        {
            string search = DetailsSearch.Text.Trim();
            if (string.IsNullOrEmpty(search))
                return true;

            return crmPluginTrace.CorrelationId.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }
  
        public bool IsStringFilterValid(ObservableCollection<IFilterProperty> list, string name)
        {
            return list.Count(e => e.IsSelected) != 0 &&
                   list.Where(e => e.IsSelected).Select(e => e.Name).Contains(name);
        }

        private void FilterEntity_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            GridFilters.SetSelectAll(sender, FilterEntities);

            FilterTraces();
        }

        private void FilterMessage_PropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            GridFilters.SetSelectAll(sender, FilterMessages);

            FilterTraces();
        }

        private void FilterMode_PropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            GridFilters.SetSelectAll(sender, FilterModes);

            FilterTraces();
        }

        private void FilterTypeName_PropertyChanged(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            GridFilters.SetSelectAll(sender, FilterTypeNames);

            FilterTraces();
        }

        private void FilterByEntity_Click(object sender, RoutedEventArgs e)
        {
            FilterEntityPopup.OpenFilterList(sender);
        }

        private void FilterByMessage_Click(object sender, RoutedEventArgs e)
        {
            FilterMessagePopup.OpenFilterList(sender);
        }

        private void FilterByMode_Click(object sender, RoutedEventArgs e)
        {
            FilterModePopup.OpenFilterList(sender);
        }
        private void FilterByTypeName_Click(object sender, RoutedEventArgs e)
        {
            FilterTypeNamePopup.OpenFilterList(sender);
        }

        private void ClearFilters_Click(object sender, RoutedEventArgs e)
        {
            if (FilterEntities[0].IsSelected != true)
                FilterEntities[0].IsSelected = true;
            if (FilterMessages[0].IsSelected != true)
                FilterMessages[0].IsSelected = true;
            if (FilterModes[0].IsSelected != true)
                FilterModes[0].IsSelected = true;
            if (FilterTypeNames[0].IsSelected != true)
                FilterTypeNames[0].IsSelected = true;

            DetailsSearch.Text = string.Empty;
        }

        private void DetailsSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterTraces();
        }

        private void TextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            textBox.SelectAll();
        }
    }
}