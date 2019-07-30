using System.Collections.ObjectModel;

namespace D365DeveloperExtensions.Core.DataGrid
{
    public class GridFilters
    {
        public static void SetSelectAll<T1, T2>(T1 sender, ObservableCollection<T2> list)
        {
            var changedFilter = (IFilterProperty)sender;
            var selectedValue = changedFilter.IsSelected;

            //Set select/unselect all
            if (string.IsNullOrEmpty(changedFilter.Value))
            {
                var allValue = changedFilter.IsSelected;
                if (allValue)
                {
                    for (var i = 1; i < list.Count; i++)
                    {
                        var filter = (IFilterProperty)list[i];
                        if (filter.IsSelected != true)
                            filter.IsSelected = true;
                    }
                }
            }
            else
            {
                var allFilter = (IFilterProperty)list[0];
                var matchCount = 0;
                for (var i = 1; i < list.Count; i++)
                {
                    var filter = (IFilterProperty)list[i];
                    if (filter.IsSelected == selectedValue)
                        matchCount++;
                }

                if (matchCount == list.Count - 1)
                {
                    if (allFilter.IsSelected != true)
                        allFilter.IsSelected = true;
                }
                else
                {
                    if (allFilter.IsSelected)
                        allFilter.IsSelected = false;
                }
            }
        }
    }
}