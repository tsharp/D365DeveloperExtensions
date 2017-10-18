using PluginTraceViewer.ViewModels;
using System.Collections.ObjectModel;

namespace PluginTraceViewer
{
    public class GridFilters
    {
        public static void SetSelectAll<T1, T2>(T1 sender, ObservableCollection<T2> list)
        {
            IFilterProperty changedFilter = (IFilterProperty)sender;
            bool selectedValue = changedFilter.IsSelected;

            //Set select/unselect all
            if (string.IsNullOrEmpty(changedFilter.Value))
            {
                bool allValue = changedFilter.IsSelected;
                if (allValue)
                {
                    for (int i = 1; i < list.Count; i++)
                    {
                        IFilterProperty filter = (IFilterProperty)list[i];
                        if (filter.IsSelected != true)
                            filter.IsSelected = true;
                    }
                }
            }
            else
            {
                IFilterProperty allFilter = (IFilterProperty)list[0];
                int matchCount = 0;
                for (int i = 1; i < list.Count; i++)
                {
                    IFilterProperty filter = (IFilterProperty)list[i];
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