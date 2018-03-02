using CrmIntellisense.Crm;
using CrmIntellisense.Models;
using D365DeveloperExtensions.Core;
using EnvDTE;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;

namespace CrmIntellisense
{
    internal class CrmCSharpCompletionSource : ICompletionSource
    {
        private readonly CrmCSharpCompletionSourceProvider _mSourceProvider;
        private readonly ITextBuffer _mTextBuffer;
        private bool _mIsDisposed;

        public CrmCSharpCompletionSource(CrmCSharpCompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            _mSourceProvider = sourceProvider;
            _mTextBuffer = textBuffer;
        }

        void ICompletionSource.AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            if (!PositionHelper.IsStringLiteral(session, _mTextBuffer))
                return;

            if (!(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(typeof(DTE)) is DTE dte))
                return;

            List<Completion> metadata = (List<Completion>)SharedGlobals.GetGlobal("CrmMetadata", dte);

            if (metadata == null)
            {
                if (CrmMetadata.Metadata == null)
                    return;

                List<CompletionValue> strList = CrmMetadata.Metadata;
                metadata = new List<Completion>();
                foreach (CompletionValue completionValue in strList)
                {
                    metadata.Add(new Completion(completionValue.Name, completionValue.Replacement,
                        completionValue.Description, MonikerHelper.GetImage(completionValue.MetadataType), null));
                }

                SharedGlobals.SetGlobal("CrmMetadata", metadata, dte);
            }

            completionSets.Add(new CompletionSet(
                "CRM",
                "CRM",
                FindTokenSpanAtPosition(session.GetTriggerPoint(_mTextBuffer), session),
                metadata,
                null)
            );
        }

        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            if (point == null)
                throw new ArgumentNullException(nameof(point));

            SnapshotPoint currentPoint = session.TextView.Caret.Position.BufferPosition - 1;
            ITextStructureNavigator navigator = _mSourceProvider.NavigatorService.GetTextStructureNavigator(_mTextBuffer);
            TextExtent extent = navigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }

        public void Dispose()
        {
            if (_mIsDisposed)
                return;

            GC.SuppressFinalize(this);
            _mIsDisposed = true;
        }
    }
}