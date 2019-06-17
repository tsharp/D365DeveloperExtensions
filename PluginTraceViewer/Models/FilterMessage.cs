using PluginTraceViewer.Resources;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginTraceViewer.Models
{
    public class FilterMessage : FilterBase
    {
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
                Name = Resource.FilterEntity_Select_All,
                Value = string.Empty,
                IsSelected = true
            });

            return filterMessages;
        }

        public static ObservableCollection<FilterMessage> ResetFilter(ObservableCollection<FilterMessage> filterMessages)
        {
            if (filterMessages[0].IsSelected != true)
                filterMessages[0].IsSelected = true;

            return filterMessages;
        }
    }
}