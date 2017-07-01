using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace CrmIntellisense
{
    [Export(typeof(ICompletionSourceProvider))]
    [ContentType("CSharp")]
    [Name("CRM CSharp Token Completion")]
    internal class CrmCSharpCompletionSourceProvider : ICompletionSourceProvider
    {
        [Import]
        internal ITextStructureNavigatorSelectorService NavigatorService { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            return new CrmCompletionSource(this, textBuffer);
        }
    }
}