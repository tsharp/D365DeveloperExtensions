using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer
{
    public class WebResourceItemHandler
    {
        public static ObservableCollection<WebResourceItem> SetDescriptions(ObservableCollection<WebResourceItem> webResourceItems,
            Guid webResourceId, string description)
        {
            foreach (var item in webResourceItems.Where(w => w.WebResourceId == webResourceId))
            {
                item.Description = description;
                item.PreviousDescription = description;
            }

            return webResourceItems;
        }

        public static ObservableCollection<WebResourceItem> ResetDescriptions(ObservableCollection<WebResourceItem> webResourceItems,
            WebResourceItem webResourceItem)
        {
            webResourceItem.PreviousDescription = webResourceItem.Description;

            foreach (var item in webResourceItems.Where(w => w.WebResourceId == webResourceItem.WebResourceId))
            {
                item.Description = item.PreviousDescription;
            }

            return webResourceItems;
        }

        public static WebResourceItem WebResourceItemFromCmdParam(object sender, ObservableCollection<WebResourceItem> webResourceItems)
        {
            var button = (Button)sender;
            var webResourceId = (Guid)button.CommandParameter;
            var webResourceItem = webResourceItems.FirstOrDefault(w => w.WebResourceId == webResourceId);

            return webResourceItem;
        }

        public static WebResourceItem SetDescriptionFromInput(WebResourceItem webResourceItem, string newDescription)
        {
            if (webResourceItem.PreviousDescription == newDescription)
                return webResourceItem;

            webResourceItem.Description = newDescription;
            webResourceItem.PreviousDescription = newDescription;

            return webResourceItem;
        }

        public static ObservableCollection<WebResourceItem> ResetPublishValues(IEnumerable<WebResourceItem> toPublish,
            ObservableCollection<WebResourceItem> webResourceItems)
        {
            var toUpdate = webResourceItems.Where(w => toPublish.Any(t => w.WebResourceId == t.WebResourceId));
            foreach (var item in toUpdate)
            {
                item.Publish = true;
            }

            return webResourceItems;
        }
    }
}