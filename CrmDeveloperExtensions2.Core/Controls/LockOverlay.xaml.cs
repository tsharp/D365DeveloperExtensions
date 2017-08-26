using CrmDeveloperExtensions2.Core.Resources;
using System.Windows;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class LockOverlay
    {
        public LockOverlay()
        {
            InitializeComponent();
        }

        public void Show(string message = null)
        {
            if (string.IsNullOrEmpty(message))
                message = Resource.LockMessage_Label_DefaultContent;

            Overlay.Visibility = Visibility.Visible;
            LockMessage.Content = message;
        }

        public void Hide()
        {
            Overlay.Visibility = Visibility.Hidden;
        }
    }
}