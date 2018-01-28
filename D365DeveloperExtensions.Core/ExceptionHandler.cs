using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.Xrm.Tooling.CrmConnectControl;
using NLog;
using System;

namespace D365DeveloperExtensions.Core
{
    public class ExceptionHandler
    {
        public static void LogException(Logger logger, string message, Exception ex)
        {
            string output = FormatExceptionOutput(message, ex);

            ExtensionLogger.LogToFile(logger, output, LogLevel.Error);

            OutputLogger.WriteToOutputWindow(output, MessageType.Error);
        }

        public static void LogCrmConnectionError(Logger logger, string message, CrmConnectionManager crmConnectionManager)
        {
            string output = FormatCrmConnectionErrorOutput(crmConnectionManager);

            ExtensionLogger.LogToFile(logger, output, LogLevel.Error);

            OutputLogger.WriteToOutputWindow(output, MessageType.Error);
        }

        public static void LogProcessError(Logger logger, string message, string errorDataReceived)
        {
            string output = FormatProcessErrorOutput(message, errorDataReceived);

            ExtensionLogger.LogToFile(logger, output, LogLevel.Error);

            OutputLogger.WriteToOutputWindow(output, MessageType.Error);
        }

        private static string FormatExceptionOutput(string message, Exception ex)
        {
            string result = $"{message}: {ex.Message}";

            if (ex.StackTrace != null)
                result += Environment.NewLine + ex.StackTrace;

            return result;
        }

        private static string FormatCrmConnectionErrorOutput(CrmConnectionManager crmConnectionManager)
        {
            return Resource.ErrorMessage_ErrorConnecting_LastError + ": " +
                   crmConnectionManager.LastError + Environment.NewLine +
                   Resource.ErrorMessage_ErrorConnecting_LastException + ": " +
                   crmConnectionManager.LastException.Message +
                   Environment.NewLine +
                   crmConnectionManager.LastException.StackTrace;
        }

        private static string FormatProcessErrorOutput(string message, string errorDataReceived)
        {
            return $"{message}: {errorDataReceived}";
        }
    }
}