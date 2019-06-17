using System;
using System.ComponentModel;
using System.Resources;

namespace D365DeveloperExtensions.Core.Localization
{
    public class LocalizedDisplayNameAttribute : DisplayNameAttribute
    {
        private readonly ResourceManager _resourceManager;
        private readonly string _resourceKey;

        public LocalizedDisplayNameAttribute(string resourceKey, Type resourceType)
        {
            _resourceManager = new ResourceManager(resourceType);
            _resourceKey = resourceKey;
        }

        public override string DisplayName
        {
            get
            {
                var displayName = _resourceManager.GetString(_resourceKey);
                return string.IsNullOrWhiteSpace(displayName) ? $@"[[{_resourceKey}]]" : displayName;
            }
        }
    }
}