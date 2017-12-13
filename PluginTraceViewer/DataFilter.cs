using CrmDeveloperExtensions2.Core.DataGrid;
using CrmDeveloperExtensions2.Core.ExtensionMethods;
using PluginTraceViewer.Models;
using PluginTraceViewer.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace PluginTraceViewer
{
    public class DataFilter
    {
        public static bool FilterItems(FilterCriteria filterCriteria)
        {
            return EntityCondition(filterCriteria.CrmPluginTrace, filterCriteria.FilterEntities)
                   && MessageCondition(filterCriteria.CrmPluginTrace, filterCriteria.FilterMessages)
                   && ModeCondition(filterCriteria.CrmPluginTrace, filterCriteria.FilterModes)
                   && (StringContainsCondition(filterCriteria.CrmPluginTrace.Details, filterCriteria.SearchText) ||
                       StringContainsCondition(filterCriteria.CrmPluginTrace.CorrelationId, filterCriteria.SearchText))
                   && TypeNameCondition(filterCriteria.CrmPluginTrace, filterCriteria.FilterTypeNames);
        }

        private static bool EntityCondition(CrmPluginTrace crmPluginTrace, ObservableCollection<FilterEntity> filterEntities)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterEntities), crmPluginTrace.Entity);
        }

        private static bool MessageCondition(CrmPluginTrace crmPluginTrace, ObservableCollection<FilterMessage> filterMessages)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterMessages), crmPluginTrace.MessageName);
        }

        private static bool ModeCondition(CrmPluginTrace crmPluginTrace, ObservableCollection<FilterMode> filterModes)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterModes), crmPluginTrace.Mode);
        }

        private static bool TypeNameCondition(CrmPluginTrace crmPluginTrace, ObservableCollection<FilterTypeName> filterTypeNames)
        {
            return IsStringFilterValid(new ObservableCollection<IFilterProperty>(filterTypeNames), crmPluginTrace.TypeName);
        }

        private static bool StringContainsCondition(string value, string search)
        {
            if (string.IsNullOrEmpty(value))
                return false;

            if (string.IsNullOrEmpty(search))
                return true;

            if (value.Contains(search, StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static bool IsStringFilterValid(ObservableCollection<IFilterProperty> list, string name)
        {
            return list.Count(e => e.IsSelected) != 0 &&
                   list.Where(e => e.IsSelected).Select(e => e.Name).Contains(name);
        }
    }
}