using D365DeveloperExtensions.Core.DataGrid;
using D365DeveloperExtensions.Core.Models;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace WebResourceDeployer.Models
{
    public class FilterTypeName : INotifyPropertyChanged, IFilterProperty
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

        public static ObservableCollection<FilterTypeName> CreateFilterList(int majorVersion)
        {
            ObservableCollection<FilterTypeName> filterTypeNames = new ObservableCollection<FilterTypeName>();

            foreach (WebResourceType webResourceType in WebResourceTypes.GetTypes(majorVersion, false))
            {
                filterTypeNames.Add(new FilterTypeName
                {
                    Name = webResourceType.Name,
                    Value = webResourceType.Name,
                    IsSelected = true
                });
            }

            filterTypeNames = new ObservableCollection<FilterTypeName>(filterTypeNames.OrderBy(e => e.Name));

            filterTypeNames.Insert(0, new FilterTypeName
            {
                Name = "Select All",
                Value = String.Empty,
                IsSelected = true
            });

            return filterTypeNames;
        }

        public static ObservableCollection<FilterTypeName> ResetFilter(ObservableCollection<FilterTypeName> filterTypeNames)
        {
            if (filterTypeNames[0].IsSelected != true)
                filterTypeNames[0].IsSelected = true;

            return filterTypeNames;
        }
    }
}