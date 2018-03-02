using System;
using System.Collections.ObjectModel;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Models
{
    public class FilterCriteria
    {
        public WebResourceItem WebResourceItem { get; set; }
        public string SearchText { get; set; }
        public Guid SolutionId { get; set; }
        public ObservableCollection<FilterState> FilterStates { get; set; }
        public ObservableCollection<FilterTypeName> FilterTypeNames { get; set; }
    }
}