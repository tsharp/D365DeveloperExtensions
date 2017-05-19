using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace CrmDeveloperExtensions.Core
{
    public static class WebBrowser
    {
        public static void OpenCrmPage(DTE dte, Uri crmUrl, string contentUrl)
        {
            bool useInternalBrowser = UserOptionsGrid.GetUseInternalBrowser(dte);
            Uri url = new Uri(crmUrl, contentUrl);

            if (useInternalBrowser) //Internal VS browser
                dte.ItemOperations.Navigate(url.ToString());
            else //User's default browser
                System.Diagnostics.Process.Start(url.ToString());
        }
    }
}
