using System;
using System.ComponentModel;
using System.Resources;

namespace D365DeveloperExtensions.Core.Localization
{
    public class LocalizedDisplayNameAttribute : DisplayNameAttribute
    {
        readonly ResourceManager _resourceManager;
        readonly string _resourceKey;

        public LocalizedDisplayNameAttribute(string resourceKey, Type resourceType)
        {
            _resourceManager = new ResourceManager(resourceType);
            _resourceKey = resourceKey;
        }

        public override string DisplayName
        {
            get
            {
                string displayName = _resourceManager.GetString(_resourceKey);
                return string.IsNullOrWhiteSpace(displayName) ? $@"[[{_resourceKey}]]" : displayName;
            }
        }
    }
}