using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace CrmDeveloperExtensions.Core
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
