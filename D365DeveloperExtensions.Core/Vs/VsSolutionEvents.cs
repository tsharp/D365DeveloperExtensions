using D365DeveloperExtensions.Core.Connection;
using D365DeveloperExtensions.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace D365DeveloperExtensions.Core.Vs
{
    public sealed class VsSolutionEvents : IVsSolutionEvents
    {
        private readonly DTE _dte;
        private readonly XrmToolingConnection _xrmToolingConnection;
        private IVsHierarchyEvents _vsHierarchyEvents;
        private uint _cookie;

        public VsSolutionEvents(DTE dte, XrmToolingConnection xrmToolingConnection)
        {
            _dte = dte;
            _xrmToolingConnection = xrmToolingConnection;
        }

        public VsSolutionEvents()
        {
        }

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            _vsHierarchyEvents = new VsHierarchyEvents(pHierarchy, _xrmToolingConnection);
            pHierarchy.AdviseHierarchyEvents(_vsHierarchyEvents, out _cookie);

            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            _vsHierarchyEvents = new VsHierarchyEvents(pHierarchy, _xrmToolingConnection);
            pHierarchy.UnadviseHierarchyEvents(_cookie);

            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        {
            foreach (Window window in _dte.Windows)
            {
                if (window.Type != vsWindowType.vsWindowTypeToolWindow)
                    continue;

                if (!HostWindow.IsD365DevExWindow(window))
                    continue;

                OutputLogger.DeleteOutputWindow();
                window.Close();
            }

            OutputLogger.DeleteOutputWindow();

            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }
    }
}