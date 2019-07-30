using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;

namespace D365DeveloperExtensions.Core
{
    public static class SharedGlobals
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        //CrmService (CrmServiceClient) = Active connection to CRM
        //UseCrmIntellisense (boolean) = Is CRM intellisense on/off

        public static object GetGlobal(string globalName, DTE dte = null)
        {
            try
            {
                if (dte == null)
                {
                    if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte2))
                        throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);
                    dte = dte2;
                }

                var globals = dte.Globals;
                return globals.VariableExists[globalName] ? globals[globalName] : null;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }

        public static void SetGlobal(string globalName, object value, DTE dte = null)
        {
            try
            {
                if (dte == null)
                {
                    if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte2))
                        throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);
                    dte = dte2;
                };

                var globals = dte.Globals;
                globals[globalName] = value;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }
    }
}