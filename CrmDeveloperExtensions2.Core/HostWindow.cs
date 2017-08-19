using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Linq;
using System.Text;

namespace CrmDeveloperExtensions2.Core
{
    public static class HostWindow
    {
        public static string SetCaption(string caption, CrmServiceClient client)
        {
            string[] parts = caption.Split('|');

            StringBuilder sb = new StringBuilder();
            sb.Append(parts[0]);

            var url = client == null ? "Not connected" : WebBrowser.GetBaseCrmUrlFomClient(client).ToString();
            if (!string.IsNullOrEmpty(url))
            {
                sb.Append(" | Connected to: ");
                sb.Append(url);
            }

            string version = client?.ConnectedOrgVersion.ToString() ?? "";
            if (!string.IsNullOrEmpty(version))
            {
                sb.Append(" | Version: ");
                sb.Append(version);
            }

            return sb.ToString();
        }

        public static bool IsCrmDevExWindow(EnvDTE.Window window)
        {
            string windowGuid = window.ObjectKind.Replace("{", String.Empty).Replace("}", String.Empty);
            return ExtensionConstants.CrmDevExWindows.Contains(windowGuid);
        }
    }
}