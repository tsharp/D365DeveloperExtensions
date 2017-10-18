using System.Windows;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class DataGridHeaderClearFilterButton
    {
        public event RoutedEventHandler Click;

        public DataGridHeaderClearFilterButton()
        {
            InitializeComponent();
        }

        protected virtual void OnClick(RoutedEventArgs e)
        {
            Click?.Invoke(this, e);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OnClick(e);
        }
    }
}