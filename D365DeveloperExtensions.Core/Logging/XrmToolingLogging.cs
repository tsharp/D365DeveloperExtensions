using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using D365DeveloperExtensions.Core.UserOptions;
using System;
using System.IO;

namespace D365DeveloperExtensions.Core.Logging
{
    public static class XrmToolingLogging
    {
        public static string GetLogFilePath()
        {
            string logFilePath = UserOptionsHelper.GetOption<string>(UserOptionProperties.XrmToolingLogFilePath);

            if (!string.IsNullOrEmpty(logFilePath))
                return Path.Combine(logFilePath, "D365DevExXrmToolingLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");

            logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            OutputLogger.WriteToOutputWindow(Resource.WarningMessage_MissingXrmToolingLogPath, MessageType.Warning);

            return Path.Combine(logFilePath, "D365DevExXrmToolingLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");
        }
    }
}