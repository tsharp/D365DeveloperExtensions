using System.Collections.Generic;

namespace TemplateWizards.Models
{
    public class CustomTemplates
    {
        public List<CustomTemplate> Templates { get; set; }
    }

    public class CustomTemplate
    {
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string FileName { get; set; }
        public string RelativePath { get; set; }
        public bool CoreReplacements { get; set; }
        public string Language { get; set; }
        public List<CustomTemplateReference> CustomTemplateReferences { get; set; }
        public List<CustomTemplateNuGetPackage> CustomTemplateNuGetPackages { get; set; }
    }

    public class CustomTemplateReference
    {
        public string Name { get; set; }
    }

    public class CustomTemplateNuGetPackage
    {
        public string Name { get; set; }
        public string Version { get; set; }
    }
}