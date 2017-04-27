using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Common;
using EnvDTE;
using Microsoft.VisualStudio.TemplateWizard;


namespace TemplateWizards
{
    public class VersionWizard : IWizard
    {
        [DllImport("mscoree.dll")]
        internal static extern int StrongNameFreeBuffer(IntPtr pbMemory);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern int StrongNameKeyGen(IntPtr wszKeyContainer, uint dwFlags, out IntPtr keyBlob, out uint keyBlobSize);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode)]
        internal static extern int StrongNameErrorInfo();

        private DTE _dte;

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            _dte = (DTE)automationObject;

            //var xamlDialog = new SdkVersionPicker
            //{
            //    HasMinimizeButton = false,
            //    HasMaximizeButton = false
            //};
            //xamlDialog.ShowModal();

            var l = new XrmToolingLogin();
            l.ShowDialog();

        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        public void RunFinished()
        {
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
        }
    }
}
