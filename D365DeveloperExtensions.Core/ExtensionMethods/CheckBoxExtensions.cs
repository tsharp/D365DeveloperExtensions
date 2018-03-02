using System.Windows.Controls;

namespace D365DeveloperExtensions.Core.ExtensionMethods
{
    public static class CheckBoxExtensions
    {
        public static bool ReturnValue(this CheckBox checkBox)
        {
            return checkBox.IsChecked.HasValue && checkBox.IsChecked.Value;
        }
    }
}