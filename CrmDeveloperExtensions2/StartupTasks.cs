using CrmDeveloperExtensions2.Core.Logging;
using EnvDTE;

//These items should run when the extension is first loaded at Visual Studio startup
namespace CrmDeveloperExtensions2
{
    public static class StartupTasks
    {
        public static void Run(DTE dte)
        {
            SetupLogging(dte);
        }

        private static void SetupLogging(DTE dte)
        {
            ExtensionLogger exLogger = new ExtensionLogger(dte);
        }
    }
}
