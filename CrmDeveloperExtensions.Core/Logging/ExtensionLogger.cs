using System;
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
        public ExtensionLogger()
        {
            CreateConfig();
        }

        private static void CreateConfig()
        {
            var config = new LoggingConfiguration();
            //TODO: User Option for log path
            var fileTarget = new FileTarget
            {
                CreateDirs = true,
                FileName = "c:\\logs\\CrmDevExLog_" + DateTime.Now.ToString("MMddyyyy") + ".log"
            };
            config.AddTarget("file", fileTarget);

            //LoggingRule rule = new LoggingRule("*", fileTarget);
            config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);

            LogManager.Configuration = config;
        }

        public static void LogToFile(DTE dte, Logger logger, string message, LogLevel logLevel)
        {         
            if (UserOptionsGrid.GetLoggingOptionBoolean(dte, "LoggingEnabled"))
                logger.Log(logLevel, message);
        }
    }
}
