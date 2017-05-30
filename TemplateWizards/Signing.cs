using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using EnvDTE;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;

namespace TemplateWizards
{
    public static class Signing
    {
        [DllImport("mscoree.dll")]
        internal static extern int StrongNameFreeBuffer(IntPtr pbMemory);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern int StrongNameKeyGen(IntPtr wszKeyContainer, uint dwFlags, out IntPtr keyBlob, out uint keyBlobSize);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode)]
        internal static extern int StrongNameErrorInfo();

        public static void GenerateKey(DTE dte, Project project, string destDirectory)
        {
            try
            {
                if (string.IsNullOrEmpty(destDirectory))
                    return;

                //Generate new key
                StatusBar.SetStatusBarValue(dte, Resources.Resource.GeneratingKeyStatusBarMessage);

                string keyFilePath = Path.Combine(destDirectory, Resources.Resource.DefaultKeyName);
                IntPtr buffer = IntPtr.Zero;

                try
                {
                    uint buffSize;
                    if (0 != StrongNameKeyGen(IntPtr.Zero, 0, out buffer, out buffSize))
                        Marshal.ThrowExceptionForHR(StrongNameErrorInfo());
                    if (buffer == IntPtr.Zero)
                        throw new InvalidOperationException();

                    var keyBuffer = new byte[buffSize];
                    Marshal.Copy(buffer, keyBuffer, 0, (int)buffSize);
                    File.WriteAllBytes(keyFilePath, keyBuffer);
                }
                finally
                {
                    StrongNameFreeBuffer(buffer);
                }

                //var props = _dte.Properties["CRM Developer Extensions", "General"];
                //string defaultKeyFileName = props.Item("DefaultProjectKeyFileName").Value;

                //foreach (ProjectItem item in project.ProjectItems)
                //{
                //    if (item.Name.ToUpper() != "MYKEY.SNK") continue;

                //    item.Name = defaultKeyFileName + ".snk";
                //    return;
                //}

                project.Properties.Item("SignAssembly").Value = "true";
                project.Properties.Item("AssemblyOriginatorKeyFile").Value = Resources.Resource.DefaultKeyName;
                project.ProjectItems.AddFromFile(keyFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Resources.Resource.GeneratingKeyFailureMessage + ":  " + ex.Message);
            }
            finally
            {
                dte.StatusBar.Clear();
            }
        }
    }
}
