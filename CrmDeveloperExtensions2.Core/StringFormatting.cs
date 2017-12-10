using System;

namespace CrmDeveloperExtensions2.Core
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