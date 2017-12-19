using System;
using CrmDeveloperExtensions2.Core.DataGrid;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public class ControlHelper
    {
        public static Popup ShowFilePopup(Popup filePopup, Grid grid)
        {
            filePopup.PlacementTarget = grid;
            filePopup.Placement = PlacementMode.Relative;
            filePopup.VerticalOffset = -3;
            filePopup.HorizontalOffset = -2;
            filePopup.IsOpen = true;

            return filePopup;
        }

        public static Popup ShowFolderPopup(Popup folderPopup, Button button)
        {
            folderPopup.PlacementTarget = button;
            folderPopup.Placement = PlacementMode.Relative;
            folderPopup.IsOpen = true;

            return folderPopup;
        }

        public static void RotateButtonImage(Button showId, DataGridLength width)
        {
            showId.RenderTransformOrigin = new Point(0.5, 0.5);
            ScaleTransform flipTrans = new ScaleTransform
            {
                ScaleX = width.UnitType == DataGridLengthUnitType.SizeToCells ? 1 : -1
            };

            showId.RenderTransform = flipTrans;
        }

        public static ComboBox ShowProjectFileList(ComboBox projectFileList, double columnWidth)
        {
            //Fixes the placement of the popup so it fits in the cell w/ the padding applied
            projectFileList.Width = columnWidth - 1;
            projectFileList.IsDropDownOpen = true;

            return projectFileList;
        }

        public static void SetPublishAll(DataGrid webResourceGrid)
        {
            CheckBox publishAll = DataGridHelpers.FindVisualChildren<CheckBox>(webResourceGrid).FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll == null)
                return;

            int allowPublish = 0;
            int publish = 0;
            foreach (WebResourceItem webResourceItem in webResourceGrid.Items)
            {
                if (webResourceItem.AllowPublish)
                    allowPublish++;
                if (webResourceItem.Publish)
                    publish++;
            }

            publishAll.IsChecked = allowPublish == publish;
        }
    }
}