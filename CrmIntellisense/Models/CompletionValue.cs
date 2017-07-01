namespace CrmIntellisense.Models
{
    public class CompletionValue
    {
        public string Name { get; set; }
        public string Replacement { get; set; }
        public string Description { get; set; }

        public CompletionValue(string name, string replacement, string description)
        {
            Name = name;
            Replacement = replacement;
            Description = description;
        }
    }
}