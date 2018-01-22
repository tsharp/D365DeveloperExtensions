using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Resources;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace CrmDeveloperExtensions2.Core
{
    public class InfoBar : Package, IVsInfoBarUIEvents
    {
        private uint _cookie;
        private readonly bool _useActiveView;
        private IVsInfoBarHost _host;
        private IVsInfoBarUIElement _element;

        public event EventHandler<IVsInfoBarUIElement> InfoBarClosed;
        public event EventHandler<InfobarActionItemEventArgs> InfobarActionItemClicked;

        public InfoBar(bool useActiveView)
        {
            _useActiveView = useActiveView;
        }

        public void HideInfoBar()
        {
            if (_host != null)
                RemoveInfoBar();
        }

        public void ShowInfoBar(InfoBarModel infoBarModel)
        {
            TryGetInfoBarData(out var host);
            _host = host;

            if (_host != null)
                CreateInfoBar(infoBarModel);
        }

        private void RemoveInfoBar()
        {
            _host.RemoveInfoBar(_element);
        }

        private void CreateInfoBar(InfoBarModel infoBarModel)
        {
            if (!(ServiceProvider.GlobalProvider.GetService(typeof(SVsInfoBarUIFactory)) is IVsInfoBarUIFactory factory))
            {
                OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_UnknownInfobarError, MessageType.Error);
                return;
            }

            IVsInfoBarUIElement element = factory.CreateInfoBar(infoBarModel);
            _element = element;
            _element.Advise(this, out _cookie);
            _host.AddInfoBar(_element);
        }

        private bool TryGetInfoBarData(out IVsInfoBarHost infoBarHost)
        {
            infoBarHost = null;

            if (_useActiveView)
            {
                // We want to get whichever window is currently in focus (including toolbars) as we could have had an exception thrown from the error list
                // or interactive window
                if (!(ServiceProvider.GlobalProvider.GetService(typeof(SVsShellMonitorSelection)) is IVsMonitorSelection monitorSelectionService) ||
                    ErrorHandler.Failed(monitorSelectionService.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_WindowFrame, out var value)))
                    return false;

                if (!(value is IVsWindowFrame frame))
                {
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_UnknownInfobarError, MessageType.Error);
                    return false;
                }

                if (ErrorHandler.Failed(frame.GetProperty((int)__VSFPROPID7.VSFPROPID_InfoBarHost, out var activeViewInfoBar)))
                    return false;

                infoBarHost = activeViewInfoBar as IVsInfoBarHost;
                return infoBarHost != null;
            }

            // Show on main window info bar
            if (!(ServiceProvider.GlobalProvider.GetService(typeof(SVsShell)) is IVsShell shell) ||
                ErrorHandler.Failed(shell.GetProperty((int)__VSSPROPID7.VSSPROPID_MainWindowInfoBarHost, out var globalInfoBar)))
                return false;

            infoBarHost = globalInfoBar as IVsInfoBarHost;
            return infoBarHost != null;
        }

        public void OnClosed(IVsInfoBarUIElement infoBarUiElement)
        {
            infoBarUiElement.Unadvise(_cookie);
            OnInfoBarClosed(infoBarUiElement);
        }

        public void OnActionItemClicked(IVsInfoBarUIElement infoBarUiElement, IVsInfoBarActionItem actionItem)
        {
            OnInfobarActionItemClicked(new InfobarActionItemEventArgs
            {
                InfoBarElement = infoBarUiElement,
                InfobarActionItem = actionItem
            });
        }

        protected virtual void OnInfoBarClosed(IVsInfoBarUIElement e)
        {
            InfoBarClosed?.Invoke(this, e);
        }

        protected virtual void OnInfobarActionItemClicked(InfobarActionItemEventArgs e)
        {
            InfobarActionItemClicked?.Invoke(this, e);
        }
    }
}