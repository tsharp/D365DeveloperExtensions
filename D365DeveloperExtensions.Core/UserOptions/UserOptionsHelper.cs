using D365DeveloperExtensions.Core.Models;

namespace D365DeveloperExtensions.Core.UserOptions
{
    public class UserOptionsHelper
    {
        public static T GetOption<T>(UserOptionProperty userOptionProperty)
        {
            var option = SharedGlobals.GetGlobal(userOptionProperty.Name);
            return (T)option;
        }

        public static void SetOption<T>(UserOptionProperty userOptionProperty, T value)
        {
            SharedGlobals.SetGlobal(userOptionProperty.Name, value);
        }
    }
}