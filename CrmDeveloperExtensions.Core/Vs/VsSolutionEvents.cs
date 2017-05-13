using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.PlatformUI.OleComponentSupport;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions.Core.Vs
{
    public sealed class VsSolutionEvents : IVsSolutionEvents
    {
        private DTE _dte;

        //IVsHierarchyEvents _vsHierarchyEvents;
        //uint _cookie;
        //private readonly WebResourceList _webResouceList;

        //public VsSolutionEvents(WebResourceList webResouceList)
        public VsSolutionEvents(DTE dte)
        {
            _dte = dte;
            //_webResouceList = webResouceList;
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
            //_vsHierarchyEvents = new VsHierarchyEvents(pHierarchy, _webResouceList);
            //pHierarchy.AdviseHierarchyEvents(_vsHierarchyEvents, out _cookie);
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
                    window.Close();
                }
            }

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
