using PluginTraceViewer.ViewModels;
using System.Collections.ObjectModel;

namespace PluginTraceViewer.Models
{
    public class FilterCriteria
    {
        public CrmPluginTrace CrmPluginTrace { get; set; }
        public string SearchText { get; set; }
        public ObservableCollection<FilterEntity> FilterEntities { get; set; }
        public ObservableCollection<FilterMessage> FilterMessages { get; set; }
        public ObservableCollection<FilterMode> FilterModes { get; set; }
        public ObservableCollection<FilterTypeName> FilterTypeNames { get; set; }
    }
}