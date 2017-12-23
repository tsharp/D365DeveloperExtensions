using CrmDeveloperExtensions2.Core.Resources;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using System;

namespace CrmDeveloperExtensions2.Core.Logging
{
    public static class OutputLogger
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void DeleteOutputWindow()
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            outWindow?.DeletePane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid);
        }

        public static void WriteToOutputWindow(string message, Enums.MessageType messageType)
        {
            message = $"{messageType.ToString().ToUpper()} | {DateTime.Now:G}: {message}";
            message = Environment.NewLine + message;

            var guidPane = VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            if (!(Package.GetGlobalService(typeof(SVsOutputWindow)) is IVsOutputWindow outputWindow))
            {
                Logger.Error(Resource.ErrorMessage_ErrorGettingOutputWindow);
                return;
            }

            if (guidPane == VSConstants.OutputWindowPaneGuid.GeneralPane_guid)
                outputWindow.CreatePane(guidPane, Resources.Resource.OutputLoggerWindowTitle, 1, 1);

            outputWindow.GetPane(guidPane, out var outputWindowPane);

            if (outputWindowPane == null) 
                return;

            outputWindowPane.Activate();
            outputWindowPane.OutputString(message);
        }
    }
}