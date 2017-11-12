using System.Collections.Generic;

namespace TemplateWizards
{
    public class ProjectDataHandler
    {
        public static void AddOrUpdateReplacements(string key, string value, ref Dictionary<string, string> replacementsDictionary)
        {
            if (replacementsDictionary.ContainsKey(key))
                replacementsDictionary[key] = value;
            else
                replacementsDictionary.Add(key, value);
        }
    }
}