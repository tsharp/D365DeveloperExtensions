using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

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
