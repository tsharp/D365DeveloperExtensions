using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.UserOptions;
using EnvDTE;

//These items should run when the extension is first loaded at Visual Studio startup
namespace CrmDeveloperExtensions2
{
    public static class StartupTasks
    {
        public static void Run(DTE dte)
        {
            SetupUserOptionsHelper(dte);           
            SetupLogging(dte);
            SetupStatusBar(dte);          
        }

        private static void SetupLogging(DTE dte)
        {
            ExtensionLogger logger = new ExtensionLogger(dte);
        }

        private static void SetupStatusBar(DTE dte)
        {
            Core.StatusBar statusBar = new Core.StatusBar(dte);
        }

        private static void SetupUserOptionsHelper(DTE dte)
        {
            UserOptionsHelper userOptionsHelper = new UserOptionsHelper(dte);
        }
    }
}