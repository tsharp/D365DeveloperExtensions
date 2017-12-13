using PluginTraceViewer.Resources;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginTraceViewer.Models
{
    public class FilterMode : FilterBase
    {
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
                Name = Resource.FilterEntity_Select_All,
                Value = String.Empty,
                IsSelected = true
            });

            return filterModes;
        }
    }
}