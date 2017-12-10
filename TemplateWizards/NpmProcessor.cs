using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using Process = System.Diagnostics.Process;
using StatusBar = CrmDeveloperExtensions2.Core.StatusBar;

namespace TemplateWizards
{
    public class NpmProcessor
    {
        public static void InstallPackage(string package, string version, string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(version))
                    version = $"@{version}";

                StatusBar.SetStatusBarValue($"{Resources.Resource.NpmPackageInstallingStatusBarMessage}: {package}{version}");

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "cmd",
                    RedirectStandardInput = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    WorkingDirectory = path,
                    UseShellExecute = false
                };
                var process = Process.Start(processStartInfo);
                if (process == null)
                {
                    MessageBox.Show($"{Resources.Resource.NpmPackageInstallFailureMessage}");
                    return;
                }

                process.StandardInput.WriteLine($"npm install --save {package}{version}");
                process.StandardInput.Flush();
                process.StandardInput.Close();
                process.WaitForExit();

                if (process.ExitCode != 0)
                    MessageBox.Show($"{Resources.Resource.NpmPackageInstallFailureMessage}: {process.StandardError.ReadToEnd()}");
            }
            finally
            {
                StatusBar.ClearStatusBarValue();
            }
        }

        public static NpmHistory GetPackageHistory(string package)
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "cmd",
                RedirectStandardInput = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                UseShellExecute = false
            };
            var process = Process.Start(processStartInfo);
            if (process == null)
            {
                MessageBox.Show($"{Resources.Resource.NpmPackageInstallFailureMessage}");
                return null;
            }

            process.StandardInput.WriteLine($"npm view {package}");
            process.StandardInput.Flush();
            process.StandardInput.Close();
            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            Regex regEx = new Regex(@"\{(.|\s)*\}");
            var m = regEx.Match(output);

            string json = m.Value;

            NpmHistory history = JsonConvert.DeserializeObject<NpmHistory>(json);

            return history;
        }
    }
}