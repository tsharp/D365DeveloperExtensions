using System.Windows.Controls;

namespace CrmDeveloperExtensions2.Core.ExtensionMethods
{
    public static class CheckBoxExtensions
    {
        public static bool ReturnValue(this CheckBox checkBox)
        {
            return checkBox.IsChecked.HasValue && checkBox.IsChecked.Value;
        }
    }
}