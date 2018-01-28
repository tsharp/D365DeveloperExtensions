using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;

namespace D365DeveloperExtensions.Core.Controls
{
    public partial class DataGridHeaderFilterPopup
    {
        public string BindingPath { get; set; }

        public DataGridHeaderFilterPopup()
        {
            InitializeComponent();         
        }

        public void OpenFilterList(object sender)
        {
            TypeNameList.SetBinding(ItemsControl.ItemsSourceProperty, new Binding(BindingPath));

            DataGridHeaderFilterButton button = (DataGridHeaderFilterButton)sender;
            FilterPopup.PlacementTarget = button;
            FilterPopup.Placement = PlacementMode.Relative;
            FilterPopup.IsOpen = true;
        }
    }
}