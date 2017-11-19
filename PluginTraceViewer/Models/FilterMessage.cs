using CrmDeveloperExtensions2.Core.DataGrid;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace PluginTraceViewer.Models
{
    public class FilterMessage : INotifyPropertyChanged, IFilterProperty
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

        public static ObservableCollection<FilterMessage> CreateFilterList(ObservableCollection<CrmPluginTrace> traces)
        {
            ObservableCollection<FilterMessage> filterMessages = new ObservableCollection<FilterMessage>(traces.GroupBy(t => t.MessageName).Select(x =>
                new FilterMessage
                {
                    Name = x.Key,
                    Value = x.Key,
                    IsSelected = true
                }).ToList());

            filterMessages = new ObservableCollection<FilterMessage>(filterMessages.OrderBy(e => e.Name));

            filterMessages.Insert(0, new FilterMessage
            {
                Name = "Select All",
                Value = String.Empty,
                IsSelected = true
            });

            return filterMessages;
        }
    }
}