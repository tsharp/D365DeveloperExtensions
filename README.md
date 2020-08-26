# News Update
It has come to my attention that Jason Lattimer has either made the original D365DeveloperExtensions project private or deleted the repository completely. I fully intend to retain this repository indefiniately and will accept pull requests from the community and attempt to maintain this code base. If you have any questions please reach out to me.

# D365 Developer Extensions
A Visual Studio extension to assist with day to day development for Dynamics CRM/365 Customer Engagement versions 2011+.  
Visual Studio 2015 & 2017

This is the new version of [CRM Developer Extensions](https://marketplace.visualstudio.com/items?itemName=tsharp.D365DeveloperExtensions)

Wondering what changed? Check out the [Change Log](https://github.com/tsharp/D365DeveloperExtensions/wiki/Change-Log).

I recommend that if using with VS 2015, uninstall the older release of CRM Developer Extensions just in case.  

Build Status: 

- ![D365DeveloperExtensions.CI](https://github.com/tsharp/D365DeveloperExtensions/workflows/D365DeveloperExtensions.CI/badge.svg)
- [![Build Status](https://orbitalforge.visualstudio.com/D365DeveloperExtensions/_apis/build/status/tsharp.D365DeveloperExtensions?branchName=master)](https://orbitalforge.visualstudio.com/D365DeveloperExtensions/_build/latest?definitionId=1&branchName=master)

**What's it's got currently**  

**Numerous project & item templates**
- Plug-ins
- Custom workflows
- Web resources
- TypeScript
- Testing
- Build your own

**Web resource deployer**
- Manage mappings between D365 organizations and Visual Studio project files   
- Publish single items or multiple items simultaneously
- Filter by solution, web resource type & managed/unmanaged
- Download web resources from D365 to your project
- Open D365 to view web resources
- Compare local version of mapped files with the D365 copy
- Add new web resources from a project file
- TypeScript friendly
- Compatible with Scott Durow's [Spkl](https://github.com/scottdurow/SparkleXrm/wiki/spkl) deployment framework

**Plug-in deployer & registration**  
- 1 click deploy plug-ins & custom workflows from Visual Studio without the SDK plug-in registration tool
- Integrated ILMerge
- Compatible with Scott Durow's [Spkl](https://github.com/scottdurow/SparkleXrm/wiki/spkl) deployment framework which allows defining registration details in code

**Solution packager UI**
- 1 click download and extraction of solution to a Visual Studio project
- Re-package solution files from a project and import to D365 

**Plug-in trace log viewer**
- Easily view and search the Plug-in Trace Logs
- Ability to delete logs

**Custom intellisense**
- Custom autocomplete for entity and attribute names from your organization

For additional details, see the [wiki](https://github.com/jlattimer/D365DeveloperExtensions/wiki) page

Post any bugs, ideas, or thoughts in the [Issues](https://github.com/jlattimer/D365DeveloperExtensions/issues) area.
