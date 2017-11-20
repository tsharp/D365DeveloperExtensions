using CrmDeveloperExtensions2.Core.DataGrid;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace PluginTraceViewer.Models
{
    public class FilterMode : INotifyPropertyChanged, IFilterProperty
    {
        private bool _isSelected;

        public string Name { get; set; }
        public string Value { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public static ObservableCollection<FilterMode> CreateFilterList(ObservableCollection<CrmPluginTrace> traces)
        {
            ObservableCollection<FilterMode> filterModes = new ObservableCollection<FilterMode>(traces.GroupBy(t => t.Mode).Select(x =>
                new FilterMode
                {
                    Name = x.Key,
                    Value = x.Key,
                    IsSelected = true
                }).ToList());

            filterModes = new ObservableCollection<FilterMode>(filterModes.OrderBy(e => e.Name));

            filterModes.Insert(0, new FilterMode
            {
                Name = "Select All",
                Value = String.Empty,
                IsSelected = true
            });

            return filterModes;
        }
    }
}