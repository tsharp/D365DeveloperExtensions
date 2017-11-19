using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using CrmDeveloperExtensions2.Core.DataGrid;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Models
{
    public class FilterState : INotifyPropertyChanged, IFilterProperty
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

        public static ObservableCollection<FilterState> CreateFilterList(ObservableCollection<WebResourceItem> webResourceItems)
        {
            ObservableCollection<FilterState> filterStates = new ObservableCollection<FilterState>(webResourceItems.GroupBy(t => t.State).Select(x =>
                new FilterState
                {
                    Name = x.Key,
                    Value = x.Key
                }).ToList());

            filterStates = new ObservableCollection<FilterState>(filterStates.OrderBy(e => e.Name));

            filterStates.Insert(0, new FilterState
            {
                Name = "Select All",
                Value = String.Empty
            });

            filterStates.First(f => f.Value == "Unmanaged").IsSelected = true;

            return filterStates;
        }
    }
}