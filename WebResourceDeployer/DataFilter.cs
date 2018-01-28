using D365DeveloperExtensions.Core.DataGrid;
using D365DeveloperExtensions.Core.ExtensionMethods;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using WebResourceDeployer.Models;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public class DataFilter
    {
        public static bool FilterItems(FilterCriteria filterCriteria)
        {
            return TypeNameCondition(filterCriteria.WebResourceItem, filterCriteria.FilterTypeNames)
                   && StateCondition(filterCriteria.WebResourceItem, filterCriteria.FilterStates)
                   && SolutionCondition(filterCriteria.WebResourceItem, filterCriteria.SolutionId)
                   && SearchNameCondition(filterCriteria.WebResourceItem, filterCriteria.SearchText);
        }

        private static bool SearchNameCondition(WebResourceItem webResourceItem, string search)
        {
            if (search.Length < 2)
                return true;

            if (!string.IsNullOrEmpty(webResourceItem.Name))
            {
                if (webResourceItem.Name.Contains(search, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            if (!string.IsNullOrEmpty(webResourceItem.DisplayName))
            {
                if (webResourceItem.DisplayName.Contains(search, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            if (!string.IsNullOrEmpty(webResourceItem.Description))
            {
                if (webResourceItem.Description.Contains(search, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static bool SolutionCondition(WebResourceItem webResourceItem, Guid solutionId)
        {
            return webResourceItem.SolutionId == solutionId;
        }

        private static bool TypeNameCondition(WebResourceItem webResourceItem, ObservableCollection<FilterTypeName> filterTypeNames)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterTypeNames), webResourceItem.TypeName);
        }

        private static bool StateCondition(WebResourceItem webResourceItem, ObservableCollection<FilterState> filterStates)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterStates), webResourceItem.State);
        }

        private static bool IsStringFilterValid(ObservableCollection<IFilterProperty> list, string name)
        {
            return list.Count(e => e.IsSelected) != 0 &&
                   list.Where(e => e.IsSelected).Select(e => e.Name).Contains(name);
        }
    }
}