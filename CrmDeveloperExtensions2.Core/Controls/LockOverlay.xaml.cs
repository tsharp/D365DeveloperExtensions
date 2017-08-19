using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using CrmDeveloperExtensions2.Core.Resources;
using System.Globalization;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class LockOverlay : UserControl
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
