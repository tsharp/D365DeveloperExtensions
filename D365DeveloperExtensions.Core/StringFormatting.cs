namespace D365DeveloperExtensions.Core
{
    public class StringFormatting
    {
        public static string RemoveBracesToUpper(string value)
        {
            value = value.Replace("{", string.Empty).Replace("}", string.Empty);

            return value.ToUpper();
        }
    }
}