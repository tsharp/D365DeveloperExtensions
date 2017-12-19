using PluginTraceViewer.Resources;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginTraceViewer.Models
{
    public class FilterTypeName : FilterBase
    {
        public static ObservableCollection<FilterTypeName> CreateFilterList(ObservableCollection<CrmPluginTrace> traces)
        {
            ObservableCollection<FilterTypeName> filterTypeNames = new ObservableCollection<FilterTypeName>(traces.GroupBy(t => t.TypeName).Select(x =>
                new FilterTypeName
                {
                    Name = x.Key,
                    Value = x.Key,
                    IsSelected = true
                }).ToList());

            filterTypeNames = new ObservableCollection<FilterTypeName>(filterTypeNames.OrderBy(e => e.Name));

            filterTypeNames.Insert(0, new FilterTypeName
            {
                Name = Resource.FilterEntity_Select_All,
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