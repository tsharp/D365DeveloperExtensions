using System;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions.Core.Logging
{
    public static class OutputLogger
    {
        private static Guid _windowsId = new Guid("C2A5E6D0-0EE6-498F-BE3E-FA96CAF3A928");
        private static IVsOutputWindowPane _customPane;
        static OutputLogger()
        {
            CreateOutputWindow();
        }

        private static void CreateOutputWindow()
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            string title = Resources.Resource.OutputLoggerWindowTitle;
            if (outWindow == null)
                return;

            outWindow.GetPane(ref _windowsId, out _customPane);
            if (_customPane != null) return; //Already exists

            outWindow.CreatePane(ref _windowsId, title, 1, 1);

            outWindow.GetPane(ref _windowsId, out _customPane);
        }

        public static void DeleteOutputWindow()
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            outWindow?.DeletePane(_windowsId);
        }

        public static void WriteToOutputWindow(string message, Enums.MessageType messageType)
        {
            message = $"{messageType.ToString().ToUpper()}: {DateTime.Now:MMddyyyy}  {message}";
            message = Environment.NewLine + message;

            _customPane.OutputString(message);
            _customPane.Activate();
        }
    }
}
