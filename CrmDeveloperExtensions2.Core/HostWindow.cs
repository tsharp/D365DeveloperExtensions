using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Linq;
using System.Text;
using CrmDeveloperExtensions2.Core.Models;

namespace CrmDeveloperExtensions2.Core
{
    public static class HostWindow
    {
        public static string SetCaption(string caption, CrmServiceClient client)
        {
            string[] parts = caption.Split('|');

            StringBuilder sb = new StringBuilder();
            sb.Append(parts[0]);

            var url = (client?.ConnectedOrgUniqueName == null) ? "Not connected" : WebBrowser.GetBaseCrmUrlFomClient(client).ToString();
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
            if (window.ObjectKind == null)
                return false;

            Guid windowGuid = new Guid(window.ObjectKind.Replace("{", String.Empty).Replace("}", String.Empty));

            return ExtensionConstants.CrmDevExToolWindows.Count(w => w.ToolWindowsId == windowGuid) > 0;
        }

        public static ToolWindow GetCrmDevExWindow(EnvDTE.Window window)
        {
            if (window.ObjectKind == null)
                return null;

            Guid windowGuid = new Guid(window.ObjectKind.Replace("{", String.Empty).Replace("}", String.Empty));

            return ExtensionConstants.CrmDevExToolWindows.FirstOrDefault(w => w.ToolWindowsId == windowGuid);
        }
    }
}