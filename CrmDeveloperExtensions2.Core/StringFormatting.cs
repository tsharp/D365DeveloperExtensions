using System;

namespace CrmDeveloperExtensions2.Core
{
    public class StringFormatting
    {
        public static string FormatProjectKind(string projectKind)
        {
            projectKind = projectKind.Replace("{", String.Empty).Replace("}", String.Empty);

            return projectKind.ToUpper();
        }
    }
}
