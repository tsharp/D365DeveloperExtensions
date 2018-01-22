using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.UserOptions;
using CrmIntellisense.Crm;
using CrmIntellisense.Resources;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Tooling.Connector;
using System.ComponentModel;

namespace CrmIntellisense
{
    public class CrmCompletionHandlerProviderBase
    {
        public InfoBarModel CreateMetadataInfoBar()
        {
            InfoBarTextSpan text = new InfoBarTextSpan(Resource.Infobar_RetrievingMetadata);
            InfoBarTextSpan[] spans = { text };
            InfoBarModel infoBarModel = new InfoBarModel(spans);

            return infoBarModel;
        }

        public bool IsIntellisenseEnabled(DTE dte)
        {
            bool useIntellisense = UserOptionsHelper.GetOption<bool>(UserOptionProperties.UseIntellisense);
            if (!useIntellisense)
                return false;

            var value = SharedGlobals.GetGlobal("UseCrmIntellisense", dte);

            bool? isEnabled = (bool?)value;
            return !(bool)!isEnabled;
        }

        public void GetData(CrmServiceClient client, InfoBar infoBar)
        {
            if (CrmMetadata.Metadata != null)
            {
                infoBar.HideInfoBar();
                return;
            }

            var bgw = new BackgroundWorker();
            
            bgw.DoWork += (_, __) => CrmMetadata.GetMetadata(client);

            bgw.RunWorkerCompleted += (_, __) => infoBar.HideInfoBar();

            bgw.RunWorkerAsync();
        }
    }
}