using EnvDTE;

namespace CrmDeveloperExtensions2.Core
{
    public static class SharedGlobals
    {
        public static object GetGlobal(string globalName, DTE dte)
        {
            Globals globals = dte.Globals;
            return globals.VariableExists[globalName] ? globals[globalName] : null;
        }

        public static void SetGlobal(string globalName, object value, DTE dte)
        {
            Globals globals = dte.Globals;
            globals[globalName] = value;
        }
    }
}