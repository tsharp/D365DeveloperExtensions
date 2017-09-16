using CrmDeveloperExtensions2.Core;
using EnvDTE;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;
using CrmIntellisense.Crm;
using Microsoft.Xrm.Tooling.Connector;

namespace CrmIntellisense
{
    [Export(typeof(IVsTextViewCreationListener))]
    [Name("CRM CSharp Token Completion Handler")]
    [ContentType("CSharp")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    internal class CrmCSharpCompletionHandlerProvider : Package, IVsTextViewCreationListener
    {
        [Import]
        internal IVsEditorAdaptersFactoryService AdapterService;
        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }
        [Import]
        internal SVsServiceProvider ServiceProvider { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            //This gets executed as each code file is loaded
            //1st
            if (!(GetGlobalService(typeof(DTE)) is DTE dte))
                return;

            bool useIntellisense = UserOptionsGrid.GetUseIntellisense(dte);
            if (!useIntellisense)
                return;

            if (!(SharedGlobals.GetGlobal("CrmService", dte) is CrmServiceClient client))
                return;

            ITextView textView = AdapterService.GetWpfTextView(textViewAdapter);
            if (textView == null)
                return;

            CrmCompletionCommandHandler CreateCommandHandler() => new CrmCompletionCommandHandler(textViewAdapter, textView, this);
            textView.Properties.GetOrCreateSingletonProperty(CreateCommandHandler);

            if (CrmMetadata.Metadata == null)
                CrmMetadata.GetMetadata(client);
        }
    }
}