using EnvDTE;
using System;
using System.IO;

namespace CrmDeveloperExtensions.Core.Logging
{
    public static class XrmToolingLogging
    {
        public static string GetLogFilePath(DTE dte)
        {
            string logFilePath = UserOptionsGrid.GetLoggingOptionString(dte, "XrmToolingLogFilePath");

            return Path.Combine(logFilePath, "CrmDevExXrmToolingLog_" + DateTime.Now.ToString("MMddyyyy") + ".log");
        }
    }
}
