using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using NLog;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using TemplateWizards.Resources;

namespace TemplateWizards
{
    public class NuGetCliProcessor
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool Install(string installPath, NuGetPackage selectedPackage)
        {
            OutputLogger.WriteToOutputWindow($"{Resource.Message_InstallingNuGetPackage}: {selectedPackage.Id} {selectedPackage.VersionText}", MessageType.Info);

            int timeout = 10000;
            string workingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            workingDirectory = $@"{workingDirectory}\SdkTools";

            using (Process process = new Process())
            {
                var processStartInfo = CreateProcessStartInfo(installPath, selectedPackage);
                process.StartInfo = processStartInfo;
                process.StartInfo.WorkingDirectory = workingDirectory;

                StringBuilder output = new StringBuilder();
                StringBuilder errorDataReceived = new StringBuilder();

                using (AutoResetEvent outputWaitHandle = new AutoResetEvent(false))
                {
                    using (AutoResetEvent errorWaitHandle = new AutoResetEvent(false))
                    {
                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (e.Data == null)
                                outputWaitHandle.Set();
                            else
                                output.AppendLine(e.Data);
                        };
                        process.ErrorDataReceived += (sender, e) =>
                        {
                            if (e.Data == null)
                                errorWaitHandle.Set();
                            else
                                errorDataReceived.AppendLine(e.Data);
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();

                        string message;
                        if (process.WaitForExit(timeout) && outputWaitHandle.WaitOne(timeout) && errorWaitHandle.WaitOne(timeout))
                        {
                            if (process.ExitCode == 0)
                            {
                                OutputLogger.WriteToOutputWindow($"{Resource.Message_InstalledNuGetPackage}: {selectedPackage.Id} {selectedPackage.VersionText}", MessageType.Info);
                                return true;
                            }

                            message = $"{Resource.Message_ErrorInstallingNuGetPackage}: {selectedPackage.Id} {selectedPackage.VersionText}";
                        }
                        else
                        {
                            message = $"{Resource.Message_TimeoutInstallingNuGetPackage}: {selectedPackage.Id} {selectedPackage.VersionText}";
                        }

                        ExceptionHandler.LogProcessError(Logger, message, errorDataReceived.ToString());
                        MessageBox.Show(message);
                    }
                }
            }

            return false;
        }

        private static ProcessStartInfo CreateProcessStartInfo(string installPath, NuGetPackage selectedPackage)
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/c \"nuget install {selectedPackage.Id} -Version {selectedPackage.VersionText} -NonInteractive -Verbosity quiet -OutputDirectory \"{installPath}\"\"",
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
                UseShellExecute = false
            };

            return processStartInfo;
        }
    }
}