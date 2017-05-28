using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using CrmDeveloperExtensions.Core.Connection;
using CrmDeveloperExtensions.Core.Logging;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.PlatformUI.OleComponentSupport;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions.Core.Vs
{
    public sealed class VsSolutionEvents : IVsSolutionEvents
    {
        private readonly DTE _dte;
        private readonly XrmToolingConnection _xrmToolingConnection;
        IVsHierarchyEvents _vsHierarchyEvents;
        uint _cookie;
        //private readonly WebResourceList _webResouceList;

        //public VsSolutionEvents(WebResourceList webResouceList)
        public VsSolutionEvents(DTE dte, XrmToolingConnection xrmToolingConnection)
        {
            _dte = dte;
            //_webResouceList = webResouceList;
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
            //TODO: do these need to be disposed of?

            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        {
            //TODO: make sure this is one of our windows
            foreach (Window window in _dte.Windows)
            {
                if (window.Type == vsWindowType.vsWindowTypeToolWindow)
                {
                    OutputLogger.DeleteOutputWindow();
                    window.Close();
                }
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
