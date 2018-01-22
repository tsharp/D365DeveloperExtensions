using CrmDeveloperExtensions2.Core.Resources;

namespace CrmDeveloperExtensions2.Core.Models
{
    public static class UserOptionProperties
    {
        public static UserOptionProperty UseInternalBrowser => new UserOptionProperty
        {
            Name = "UseInternalBrowser",
            Page = Resource.UserOptions_Page_WebBrowser
        };

        public static UserOptionProperty ExtensionLoggingEnabled => new UserOptionProperty
        {
            Name = "ExtensionLoggingEnabled",
            Page = Resource.UserOptions_Page_Logging
        };
        public static UserOptionProperty ExtensionLogFilePath => new UserOptionProperty
        {
            Name = "ExtensionLogFilePath",
            Page = Resource.UserOptions_Page_Logging
        };
        public static UserOptionProperty XrmToolingLoggingEnabled => new UserOptionProperty
        {
            Name = "XrmToolingLoggingEnabled",
            Page = Resource.UserOptions_Page_Logging
        };
        public static UserOptionProperty XrmToolingLogFilePath => new UserOptionProperty
        {
            Name = "XrmToolingLogFilePath",
            Page = Resource.UserOptions_Page_Logging
        };
        public static UserOptionProperty PluginRegistrationToolPath => new UserOptionProperty
        {
            Name = "PluginRegistrationToolPath",
            Page = Resource.UserOptions_Page_Tools
        };
        public static UserOptionProperty SolutionPackagerToolPath => new UserOptionProperty
        {
            Name = "SolutionPackagerToolPath",
            Page = Resource.UserOptions_Page_Tools
        };
        public static UserOptionProperty CrmSvcUtilToolPath => new UserOptionProperty
        {
            Name = "CrmSvcUtilToolPath",
            Page = Resource.UserOptions_Page_Tools
        };
        public static UserOptionProperty CustomTemplatesPath => new UserOptionProperty
        {
            Name = "CustomTemplatesPath",
            Page = Resource.UserOptions_Page_Templates
        };
        public static UserOptionProperty DefaultKeyFileName => new UserOptionProperty
        {
            Name = "DefaultKeyFileName",
            Page = Resource.UserOptions_Page_Templates
        };
        public static UserOptionProperty UseIntellisense => new UserOptionProperty
        {
            Name = "UseIntellisense",
            Page = Resource.UserOptions_Page_Intellisense
        };
        public static UserOptionProperty IntellisenseEntityTriggerCharacter => new UserOptionProperty
        {
            Name = "IntellisenseEntityTriggerCharacter",
            Page = Resource.UserOptions_Page_Intellisense
        };
        public static UserOptionProperty IntellisenseFieldTriggerCharacter => new UserOptionProperty
        {
            Name = "IntellisenseFieldTriggerCharacter",
            Page = Resource.UserOptions_Page_Intellisense
        };
    }
}