using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace PluginTraceViewer.ViewModels
{
    public class FilterEntity : INotifyPropertyChanged, IFilterProperty
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

        public static ObservableCollection<FilterEntity> CreateFilterList(ObservableCollection<CrmPluginTrace> traces)
        {
            ObservableCollection<FilterEntity> filterEntities = new ObservableCollection<FilterEntity>(traces.GroupBy(t => t.Entity).Select(x =>
                new FilterEntity
                {
                    Name = x.Key,
                    Value = x.Key,
                    IsSelected = true
                }).ToList());

            filterEntities = new ObservableCollection<FilterEntity>(filterEntities.OrderBy(e => e.Name));

            filterEntities.Insert(0, new FilterEntity
            {
                Name = "Select All",
                Value = String.Empty,
                IsSelected = true
            });

            return filterEntities;
        }
    }
}