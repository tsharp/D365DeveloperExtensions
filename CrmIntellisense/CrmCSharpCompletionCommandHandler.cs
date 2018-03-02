using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Runtime.InteropServices;

namespace CrmIntellisense
{
    internal class CrmCSharpCompletionCommandHandler : IOleCommandTarget
    {
        private readonly IOleCommandTarget _mNextCommandHandler;
        private readonly ITextView _mTextView;
        private readonly CrmCSharpCompletionHandlerProvider _mProvider;
        private ICompletionSession _mSession;
        private readonly char _entityTriggerCharacter;
        private readonly char _entityFieldCharacter;

        internal CrmCSharpCompletionCommandHandler(IVsTextView textViewAdapter, ITextView textView, CrmCSharpCompletionHandlerProvider provider)
        {
            _mTextView = textView;
            _mProvider = provider;
            _entityTriggerCharacter = char.Parse(UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseEntityTriggerCharacter));
            _entityFieldCharacter = char.Parse(UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseFieldTriggerCharacter));

            textViewAdapter.AddCommandFilter(this, out _mNextCommandHandler);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return _mNextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdId, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (VsShellUtilities.IsInAutomationFunction(_mProvider.ServiceProvider))
                return _mNextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);

            uint commandId = nCmdId;
            char typedChar = char.MinValue;

            //Make sure the input is a char before getting it
            if (pguidCmdGroup == VSConstants.VSStd2K && nCmdId == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
                typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);

            if (nCmdId == (uint)VSConstants.VSStd2KCmdID.RETURN || nCmdId == (uint)VSConstants.VSStd2KCmdID.TAB)
            {
                // Check for a a selection
                // ReSharper disable once MergeSequentialChecksWhenPossible
                if (_mSession != null && !_mSession.IsDismissed)
                {
                    // Selection is fully selected, commit the current session
                    if (_mSession.SelectedCompletionSet.SelectionStatus.IsSelected)
                    {
                        _mSession.Commit();
                        // Don't add the character to the buffer
                        return VSConstants.S_OK;
                    }

                    // No selection, dismiss the session
                    _mSession.Dismiss();
                }
            }

            // Pass along the command so the char is added to the buffer
            int retVal = _mNextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);
            bool handled = false;

            if (!typedChar.Equals(char.MinValue) && char.IsLetterOrDigit(typedChar) ||
                !typedChar.Equals(char.MinValue) && (typedChar == _entityFieldCharacter || typedChar == _entityTriggerCharacter))
            {
                // ReSharper disable once MergeSequentialChecksWhenPossible
                if (_mSession == null || _mSession.IsDismissed) // No active session, bring up completion
                {
                    TriggerCompletion();
                    _mSession?.Filter();
                }
                else    // Completion session is already active, so just filter
                    _mSession.Filter();

                handled = true;
            }
            else if (commandId == (uint)VSConstants.VSStd2KCmdID.BACKSPACE || commandId == (uint)VSConstants.VSStd2KCmdID.DELETE) // Redo the filter if there is a deletion
            {
                // ReSharper disable once MergeSequentialChecksWhenPossible
                if (_mSession != null && !_mSession.IsDismissed)
                    _mSession.Filter();

                handled = true;
            }

            return handled ? VSConstants.S_OK : retVal;
        }

        private void TriggerCompletion()
        {
            // Caret must be in a non-projection location 
            SnapshotPoint? caretPoint = _mTextView.Caret.Position.Point.GetPoint(
                textBuffer => !textBuffer.ContentType.IsOfType("projection"), PositionAffinity.Predecessor);
            if (!caretPoint.HasValue)
                return;

            _mSession = _mProvider.CompletionBroker.CreateCompletionSession(_mTextView,
                caretPoint.Value.Snapshot.CreateTrackingPoint(caretPoint.Value.Position, PointTrackingMode.Positive), true);

            // Subscribe to the Dismissed event on the session 
            _mSession.Dismissed += OnSessionDismissed;
            _mSession.Start();
        }

        private void OnSessionDismissed(object sender, EventArgs e)
        {
            _mSession.Dismissed -= OnSessionDismissed;
            _mSession = null;
        }
    }
}