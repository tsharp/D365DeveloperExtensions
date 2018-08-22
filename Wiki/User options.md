## Overview
What the various user options are intended for.

## To find user options
Visual Studio Tools -> Options -> D365 DevEx

### External Tools Options
* **CrmSvcUtil tool Path** - Path to the local folder containing _CrmSvcUtil.exe_ from the SDK.
* **Plug-in Registration tool path** - Path to the local folder containing _PluginRegistration.exe_ from the SDK.
* **Solution Packager tool path** - Path to the local folder containing _SolutionPackager.exe_ from the SDK.

### [Intellisense](https://github.com/jlattimer/D365DeveloperExtensions/wiki/8.-CRM-Intellisense) Options
* **Entity Trigger Character** - The character used to trigger entity name completion when using CRM Intellisense. Defaults to "$".
* **Field Trigger Character** - The character used to trigger attribute name completion when using CRM Intellisense. Defaults to "_".
* **Use Intellisense?** - Determines if the option to use CRM Intellisense will be displayed or not.

### Logging Options
* **Enable detailed extension logging?** - Turns on extra tracing details built into the extension.
* **Enable Xrm.Tooling logging?** - Turns on extra logging specific to the Xrm.Tooling connection that Microsoft provides.
* **Extension log file path** - Folder where extension logs are placed.
* **Xrm.Tooling log file path?** - Folder where Xrm.Tooling logs are placed.

### Template Options
* **Default key file name** - Strong keys generated for projects will be given this name by default.
* **Path to custom templates folder** - Folder containing user created custom item templates.

### Web Browser Options
* **Use internal VS web browser?** - Open web content inside Visual Studio using the built in browser or use the system's default browser.