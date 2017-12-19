using PluginTraceViewer.Resources;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginTraceViewer.Models
{
    public class FilterEntity : FilterBase
    {
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
                Name = Resource.FilterEntity_Select_All,
                Value = String.Empty,
                IsSelected = true
            });

            return filterEntities;
        }

        public static ObservableCollection<FilterEntity> ResetFilter(ObservableCollection<FilterEntity> filterEntities)
        {
            if (filterEntities[0].IsSelected != true)
                filterEntities[0].IsSelected = true;

            return filterEntities;
        }
    }
}