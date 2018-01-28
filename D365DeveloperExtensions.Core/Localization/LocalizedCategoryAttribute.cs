using System;
using System.ComponentModel;
using System.Resources;

namespace D365DeveloperExtensions.Core.Localization
{
    public class LocalizedCategoryAttribute : CategoryAttribute
    {
        private readonly ResourceManager _resourceManager;
        private readonly string _resourceKey;

        public LocalizedCategoryAttribute(string resourceKey, Type resourceType)
        {
            _resourceManager = new ResourceManager(resourceType);
            _resourceKey = resourceKey;
        }

        protected override string GetLocalizedString(string value)
        {
            string category = _resourceManager.GetString(_resourceKey);
            return string.IsNullOrWhiteSpace(category) ? $"[[{_resourceKey}]]" : category;
        }
    }
}