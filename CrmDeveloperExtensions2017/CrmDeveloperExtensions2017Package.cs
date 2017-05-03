using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using WebResourceDeployer;

namespace CrmDeveloperExtensions2017
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.0.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(WebResourceDeployerHost))]
    [Guid(PackageGuids.GuidCrmDeveloperExtensionsPkgString)]
    //[ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]
    public sealed class CrmDeveloperExtensions2017Package : Package
    {
        protected override void Initialize()
        {
            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs == null) return;

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(PackageGuids.GuidCrmDevExCmdSet, PackageIds.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand(ShowWebResourceDeployer, wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);
        }

        private void ShowWebResourceDeployer(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(WebResourceDeployerHost), 0, true);
            if (window?.Frame == null)
                throw new NotSupportedException("Cannot create tool window.");

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
