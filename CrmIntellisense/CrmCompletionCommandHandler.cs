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
    internal class CrmCompletionCommandHandler : IOleCommandTarget
    {
        private readonly IOleCommandTarget _mNextCommandHandler;
        private readonly ITextView _mTextView;
        private readonly CrmCSharpCompletionHandlerProvider _mProvider;
        private ICompletionSession _mSession;

        internal CrmCompletionCommandHandler(IVsTextView textViewAdapter, ITextView textView, CrmCSharpCompletionHandlerProvider provider)
        {
            _mTextView = textView;
            _mProvider = provider;

            //2nd
            //add the command to the command chain
            textViewAdapter.AddCommandFilter(this, out _mNextCommandHandler);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return _mNextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdId, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (VsShellUtilities.IsInAutomationFunction(_mProvider.ServiceProvider))
            {
                return _mNextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);
            }
            //make a copy of this so we can look at it after forwarding some commands
            uint commandId = nCmdId;
            char typedChar = char.MinValue;
            //make sure the input is a char before getting it
            if (pguidCmdGroup == VSConstants.VSStd2K && nCmdId == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
            {
                typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
            }

            //check for a commit character
            //if (nCmdId == (uint)VSConstants.VSStd2KCmdID.RETURN
            //	|| nCmdId == (uint)VSConstants.VSStd2KCmdID.TAB
            //	|| (char.IsWhiteSpace(typedChar) || char.IsPunctuation(typedChar)))

            if (nCmdId == (uint)VSConstants.VSStd2KCmdID.RETURN //remove space and punctuation because it might interfere with legit entries inside a string
                || nCmdId == (uint)VSConstants.VSStd2KCmdID.TAB)
            {
                //check for a a selection
                if (_mSession != null && !_mSession.IsDismissed)
                {
                    //if the selection is fully selected, commit the current session
                    if (_mSession.SelectedCompletionSet.SelectionStatus.IsSelected)
                    {
                        _mSession.Commit();
                        //also, don't add the character to the buffer
                        return VSConstants.S_OK;
                    }

                    //if there is no selection, dismiss the session
                    _mSession.Dismiss();
                }
            }

            //pass along the command so the char is added to the buffer
            int retVal = _mNextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);
            bool handled = false;
            //if (!typedChar.Equals(char.MinValue) && char.IsLetterOrDigit(typedChar))
            if ((!typedChar.Equals(char.MinValue) && char.IsLetterOrDigit(typedChar)) || (!typedChar.Equals(char.MinValue) && (typedChar == '_' || typedChar == '$')))
            {
                if (_mSession == null || _mSession.IsDismissed) // If there is no active session, bring up completion
                {
                    TriggerCompletion();
                    _mSession?.Filter();
                }
                else    //the completion session is already active, so just filter
                {
                    _mSession.Filter();
                }
                handled = true;
            }
            else if (commandId == (uint)VSConstants.VSStd2KCmdID.BACKSPACE   //redo the filter if there is a deletion
                || commandId == (uint)VSConstants.VSStd2KCmdID.DELETE)
            {
                if (_mSession != null && !_mSession.IsDismissed)
                    _mSession.Filter();
                handled = true;
            }
            return handled ? VSConstants.S_OK : retVal;
        }

        private void TriggerCompletion()
        {
            //the caret must be in a non-projection location 
            SnapshotPoint? caretPoint =
            _mTextView.Caret.Position.Point.GetPoint(
            textBuffer => (!textBuffer.ContentType.IsOfType("projection")), PositionAffinity.Predecessor);
            if (!caretPoint.HasValue)
            {
                return;
            }

            var lineText = caretPoint.Value.GetContainingLine().GetText();
            bool inside2String = false;
            bool inside1String = false;
            int i = caretPoint.Value.GetContainingLine().Start.Position;
            foreach (char c in lineText)
            {
                if (i >= caretPoint.Value.Position)
                    break;

                if (c == '"')
                    inside2String = !inside2String;

                if (c == '\'')
                    inside1String = !inside1String;
                i++;
            }

            if (!inside2String && !inside1String)
                return;

            _mSession = _mProvider.CompletionBroker.CreateCompletionSession
            (_mTextView,
                caretPoint.Value.Snapshot.CreateTrackingPoint(caretPoint.Value.Position, PointTrackingMode.Positive),
                true);

            //subscribe to the Dismissed event on the session 
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