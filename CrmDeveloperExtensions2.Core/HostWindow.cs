using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Resources;
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

            var url = client?.ConnectedOrgUniqueName == null
                ? Resource.HostWindow_SetCaption_NotConnected
                : WebBrowser.GetBaseCrmUrlFomClient(client).ToString();

            if (!string.IsNullOrEmpty(url))
            {
                sb.Append($" | {Resource.HostWindow_SetCaption_ConnectedTo}: ");

                if (url.EndsWith("/"))
                    url = url.TrimEnd('/');

                sb.Append(url);
            }

            string version = client?.ConnectedOrgVersion.ToString() ?? "";
            if (!string.IsNullOrEmpty(version))
            {
                sb.Append($" | {Resource.HostWindow_SetCaption_Version}: ");
                sb.Append(version);
            }

            return sb.ToString();
        }

        public static bool IsCrmDevExWindow(EnvDTE.Window window)
        {
            if (window.ObjectKind == null)
                return false;

            Guid windowGuid = new Guid(StringFormatting.RemoveBracesToUpper(window.ObjectKind));

            return ExtensionConstants.CrmDevExToolWindows.Count(w => w.ToolWindowsId == windowGuid) > 0;
        }

        public static ToolWindow GetCrmDevExWindow(EnvDTE.Window window)
        {
            if (window.ObjectKind == null)
                return null;

            Guid windowGuid = new Guid(StringFormatting.RemoveBracesToUpper(window.ObjectKind));

            return ExtensionConstants.CrmDevExToolWindows.FirstOrDefault(w => w.ToolWindowsId == windowGuid);
        }
    }
}