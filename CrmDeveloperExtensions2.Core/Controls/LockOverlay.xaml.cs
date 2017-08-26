using CrmDeveloperExtensions2.Core.Resources;
using EnvDTE;
using System;
using System.Windows;
using System.Windows.Threading;

namespace CrmDeveloperExtensions2.Core.Controls
{
    public partial class LockOverlay
    {
        public LockOverlay()
        {
            InitializeComponent();
        }

        private void Show(string message = null)
        {
            if (string.IsNullOrEmpty(message))
                message = Resource.LockMessage_Label_DefaultContent;

            Overlay.Visibility = Visibility.Visible;
            LockMessage.Content = message;
        }

        private void Hide()
        {
            Overlay.Visibility = Visibility.Hidden;
        }

        public void ShowMessage(DTE dte, string message, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.SetStatusBarValue(dte, "Working...", (vsStatusAnimation)animation);
                        Show(message);
                    }
                ));
        }

        public void HideMessage(DTE dte, vsStatusAnimation? animation = null)
        {
            Dispatcher.Invoke(DispatcherPriority.Normal,
                new Action(() =>
                    {
                        if (animation != null)
                            StatusBar.ClearStatusBarValue(dte, (vsStatusAnimation)animation);
                        Hide();
                    }
                ));
        }
    }
}