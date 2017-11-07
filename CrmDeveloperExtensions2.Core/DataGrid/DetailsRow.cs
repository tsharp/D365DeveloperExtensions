using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CrmDeveloperExtensions2.Core.DataGrid
{
    public static class DetailsRow
    {
        public static void ShowHideDetailsRow(object sender)
        {
            for (var vis = sender as Visual; vis != null; vis = VisualTreeHelper.GetParent(vis) as Visual)
            {
                if (!(vis is DataGridRow))
                    continue;

                var row = (DataGridRow)vis;
                row.DetailsVisibility = row.DetailsVisibility == Visibility.Visible
                    ? Visibility.Collapsed
                    : Visibility.Visible;
                break;
            }
        }

        public static T GetDataGridRowControl<T>(object sender, string controlName)
        {
            for (var vis = sender as Visual; vis != null; vis = VisualTreeHelper.GetParent(vis) as Visual)
            {
                if (!(vis is DataGridRow))
                    continue;

                var row = (DataGridRow)vis;

                Control control = DataGridHelpers.FindVisualChildren<Control>(row).FirstOrDefault(t => t.Name == controlName);
                return (T)Convert.ChangeType(control, typeof(T));
            }

            return default(T);
        }
    }
}