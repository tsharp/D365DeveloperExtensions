using System;
using System.IO;
using System.Runtime.CompilerServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell.Interop;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace CrmDeveloperExtensions.Core.Logging
{
    public class ExtensionLogger
    {
        public ExtensionLogger(DTE dte)
        {
            CreateConfig(dte);
        }

        private static void CreateConfig(DTE dte)
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                FileName = GetLogFilePath(dte)
            };
            config.AddTarget("file", fileTarget);

            config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);

            LogManager.Configuration = config;
        }

        public static void LogToFile(DTE dte, Logger logger, string message, LogLevel logLevel)
        {
            if (UserOptionsGrid.GetLoggingOptionBoolean(dte, "LoggingEnabled"))
                logger.Log(logLevel, message);
        }

        private static string GetLogFilePath(DTE dte)
        {
            string logFilePath = UserOptionsGrid.GetLoggingOptionString(dte, "LogFilePath");

            return Path.Combine(logFilePath, "CrmDevExLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");
        }
    }
}
