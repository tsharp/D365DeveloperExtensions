using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using D365DeveloperExtensions.Core.UserOptions;
using EnvDTE;
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;

namespace D365DeveloperExtensions.Core.Logging
{
    public class ExtensionLogger
    {
        private static DTE _dte;

        public ExtensionLogger(DTE dte)
        {
            _dte = dte;
            CreateConfig();
        }

        private static void CreateConfig()
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                FileName = GetLogFilePath()
            };
            config.AddTarget("file", fileTarget);

            config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);

            LogManager.Configuration = config;
        }

        public static void LogToFile(Logger logger, string message, LogLevel logLevel)
        {
            if (UserOptionsHelper.GetOption<bool>(UserOptionProperties.ExtensionLoggingEnabled))
                logger.Log(logLevel, message);
        }

        private static string GetLogFilePath()
        {
            string logFilePath = UserOptionsHelper.GetOption<string>(UserOptionProperties.ExtensionLogFilePath);
            if (!string.IsNullOrEmpty(logFilePath))
                return Path.Combine(logFilePath, CreateLogFileName());

            logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            OutputLogger.WriteToOutputWindow(Resource.WarningMessage_MissingExtensionLogPath, MessageType.Warning);

            return Path.Combine(logFilePath, CreateLogFileName());
        }

        private static string CreateLogFileName()
        {
            return $"D365DevExLog_{DateTime.Now:MMddyyyy}.log";
        }
    }
}