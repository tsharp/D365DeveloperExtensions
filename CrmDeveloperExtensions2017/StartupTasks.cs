using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrmDeveloperExtensions.Core;
using CrmDeveloperExtensions.Core.Logging;
using CrmDeveloperExtensions.Core.Vs;
using EnvDTE;

//These items should run when the extension is first loaded at Visual Studio startup
namespace CrmDeveloperExtensions2017
{
    public static class StartupTasks
    {
        public static void Run(DTE dte)
        {
            SetupLogging(dte);

            VsSolutionEvents events = new VsSolutionEvents(dte);
        }

        private static void SetupLogging(DTE dte)
        {
            ExtensionLogger exLogger = new ExtensionLogger();
       
            //NLog.ExtensionLogger logger = new NLog.ExtensionLogger();
            //SharedGlobals.SetGlobal("logger", logger, dte);      
        }
    }
}
