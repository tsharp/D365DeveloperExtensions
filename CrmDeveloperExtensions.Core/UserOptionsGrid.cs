using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.Adapters;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace CrmDeveloperExtensions.Core
{
    //Keep this CLSCompliant attribute here
    [CLSCompliant(false), ComVisible(true)]
    public class UserOptionsGrid : DialogPage
    {
        [Category("Logging Options")]
        [DisplayName("Enable detailed extension logging?")]
        [Description("Detailed extension logging will log a bunch of stuff")]
        public bool LoggingEnabled { get; set; } = false;

        [Category("Logging Options")]
        [DisplayName("Log file path")]
        [Description("Path to log file storage")]
        public string LogFilePath { get; set; } = String.Empty;

        public UserOptionsGrid()
        {
            
        }

        //protected override void OnClosed(EventArgs e)
        //{
        //    base.OnClosed(e);
        //    var g = 1;
        //}

        //protected override void OnApply(PageApplyEventArgs e)
        //{
            
        //    base.OnApply(e);
        //    var g = 1;
        //}

        public static bool GetLoggingOptionBoolean(DTE dte, string propertyName)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsLoggingPage];

            return (bool)props.Item(propertyName).Value;
        }

        public static string GetLoggingOptionString(DTE dte, string propertyName)
        {
            var props = dte.Properties[Resources.Resource.UserOptionsCategory,
                Resources.Resource.UserOptionsLoggingPage];

            return props.Item(propertyName).Value.ToString();
        }
    }
}
