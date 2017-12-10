using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.Resources;
using EnvDTE;

namespace CrmDeveloperExtensions2.Core.UserOptions
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
    }
}