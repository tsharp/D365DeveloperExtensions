using System;
using CrmDeveloperExtensions.Core.Connection;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace CrmDeveloperExtensions.Core.Vs
{
    public sealed class VsHierarchyEvents : IVsHierarchyEvents
    {
        private readonly IVsHierarchy _hierarchy;
        private readonly XrmToolingConnection _xrmToolingConnection;

        public VsHierarchyEvents(IVsHierarchy hierarchy, XrmToolingConnection xrmToolingConnection)
        {
            _hierarchy = hierarchy;
            _xrmToolingConnection = xrmToolingConnection;
        }

        int IVsHierarchyEvents.OnInvalidateIcon(IntPtr hicon)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnInvalidateItems(uint itemidParent)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded)
        {
            object itemExtObject;
            if (_hierarchy.GetProperty(itemidAdded, (int) __VSHPROPID.VSHPROPID_ExtObject, out itemExtObject) != VSConstants.S_OK)
                return VSConstants.S_OK;

            var projectItem = itemExtObject as ProjectItem;
            if (projectItem == null)
                return VSConstants.S_OK;

            Guid type = new Guid(projectItem.Kind);
            if (type == VSConstants.GUID_ItemType_PhysicalFile || type == VSConstants.GUID_ItemType_PhysicalFolder)
                _xrmToolingConnection.ProjectItemMoveAdded(projectItem);

            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemDeleted(uint itemid)
        {
            object itemExtObject;
            if (_hierarchy.GetProperty(itemid, (int) __VSHPROPID.VSHPROPID_ExtObject, out itemExtObject) !=VSConstants.S_OK)
                return VSConstants.S_OK;

            var projectItem = itemExtObject as ProjectItem;
            if (projectItem == null)
                return VSConstants.S_OK;

            Guid type = new Guid(projectItem.Kind);
            if (type == VSConstants.GUID_ItemType_PhysicalFile || type == VSConstants.GUID_ItemType_PhysicalFolder)
                _xrmToolingConnection.ProjectItemMoveDeleted(projectItem);

            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemsAppended(uint itemidParent)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnPropertyChanged(uint itemid, int propid, uint flags)
        {
            return VSConstants.S_OK;
        }
    }
}
