using EnvDTE;

namespace D365DeveloperExtensions.Core
{
    public static class SharedGlobals
    {
        //CrmService (CrmServiceClient) = Active connection to CRM
        //UseCrmIntellisense (boolean) = Is CRM intellisense on/off

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