using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Resources;
using CrmDeveloperExtensions2.Core.UserOptions;
using System;
using System.IO;

namespace CrmDeveloperExtensions2.Core.Logging
{
    public static class XrmToolingLogging
    {
        public static string GetLogFilePath()
        {
            string logFilePath = UserOptionsHelper.GetOption<string>(UserOptionProperties.XrmToolingLogFilePath);

            if (!string.IsNullOrEmpty(logFilePath))
                return Path.Combine(logFilePath, "CrmDevExXrmToolingLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");

            logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            OutputLogger.WriteToOutputWindow(Resource.WarningMessage_MissingXrmToolingLogPath, MessageType.Warning);

            return Path.Combine(logFilePath, "CrmDevExXrmToolingLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");
        }
    }
}