using System;

namespace D365DeveloperExtensions.Core
{
    public class StringFormatting
    {
        public static string RemoveBracesToUpper(string value)
        {
            value = value.Replace("{", String.Empty).Replace("}", String.Empty);

            return value.ToUpper();
        }
    }
}