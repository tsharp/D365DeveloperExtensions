using D365DeveloperExtensions.Core.Logging;

//These items should run when the extension is first loaded at Visual Studio startup
namespace D365DeveloperExtensions
{
    public static class StartupTasks
    {
        public static void Run()
        {
            SetupLogging();
        }

        private static void SetupLogging()
        {
            var logger = new ExtensionLogger();
        }
    }
}