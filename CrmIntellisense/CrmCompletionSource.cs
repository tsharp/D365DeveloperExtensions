using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using CrmIntellisense.Crm;
using CrmIntellisense.Models;

namespace CrmIntellisense
{
    internal class CrmCompletionSource : ICompletionSource
    {
        private readonly CrmCSharpCompletionSourceProvider _mSourceProvider;
        private readonly ITextBuffer _mTextBuffer;
        private List<Completion> _mCompList;

        public CrmCompletionSource(CrmCSharpCompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            _mSourceProvider = sourceProvider;
            _mTextBuffer = textBuffer;
        }

        void ICompletionSource.AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            List<CompletionValue> strList = CrmMetadata.Metadata;
            _mCompList = new List<Completion>();
            //foreach (string str in strList)
            //	_mCompList.Add(new Completion(str, str, str, null, null));
            foreach (CompletionValue completionValue in strList)
            {
                _mCompList.Add(new Completion(completionValue.Name, completionValue.Replacement, completionValue.Description, null, null));
            }

            completionSets.Add(new CompletionSet(
                "CRM",    //the non-localized title of the tab
                "CRM",    //the display title of the tab
                FindTokenSpanAtPosition(session.GetTriggerPoint(_mTextBuffer),
                session),
                _mCompList,
                null)
            );
        }
        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            if (point == null) throw new ArgumentNullException(nameof(point));
            SnapshotPoint currentPoint = (session.TextView.Caret.Position.BufferPosition) - 1;
            ITextStructureNavigator navigator = _mSourceProvider.NavigatorService.GetTextStructureNavigator(_mTextBuffer);
            TextExtent extent = navigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }

        private bool _mIsDisposed;
        public void Dispose()
        {
            if (_mIsDisposed) return;
            GC.SuppressFinalize(this);
            _mIsDisposed = true;
        }
    }
}