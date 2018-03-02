using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.Resources;
using EnvDTE;

namespace D365DeveloperExtensions.Core.UserOptions
{
    public class UserOptionsHelper
    {
        private static DTE _dte;

        public UserOptionsHelper(DTE dte)
        {
            _dte = dte;
        }

        public static T GetOption<T>(UserOptionProperty userOptionProperty)
        {
            var props = _dte.Properties[Resource.UserOptionsCategory, userOptionProperty.Page];

            return (T)props.Item(userOptionProperty.Name).Value;
        }

        public static void SetOption<T>(UserOptionProperty userOptionProperty, T value)
        {
            var props = _dte.Properties[Resource.UserOptionsCategory, userOptionProperty.Page];
            Property setting = props.Item(userOptionProperty.Name);
            setting.Value = value;
        }
    }
}