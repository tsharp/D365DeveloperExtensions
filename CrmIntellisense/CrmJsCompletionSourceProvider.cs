using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace CrmIntellisense
{
    [Export(typeof(ICompletionSourceProvider))]
    [ContentType("JavaScript")]
    [ContentType("JScript")]
    //[ContentType("htmlx")]
    [ContentType("TypeScript")]
    [Name("CRM JS Token Completion")]
    internal class CrmJsCompletionSourceProvider : ICompletionSourceProvider
    {
        [Import]
        internal ITextStructureNavigatorSelectorService NavigatorService { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            return new CrmJsCompletionSource(this, textBuffer);
        }
    }
}