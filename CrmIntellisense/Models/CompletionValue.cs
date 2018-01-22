using CrmDeveloperExtensions2.Core.Enums;

namespace CrmIntellisense.Models
{
    public class CompletionValue
    {
        public string Name { get; set; }
        public string Replacement { get; set; }
        public string Description { get; set; }
        public MetadataType MetadataType { get; set; }

        public CompletionValue(string name, string replacement, string description, MetadataType metadataType)
        {
            Name = name;
            Replacement = replacement;
            Description = description;
            MetadataType = metadataType;
        }
    }
}