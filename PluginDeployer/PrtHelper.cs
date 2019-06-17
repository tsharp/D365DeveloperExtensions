using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using NLog;
using PluginDeployer.Resources;
using System;
using System.IO;
using System.Windows;

namespace PluginDeployer
{
    public class PrtHelper
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void OpenPrt()
        {
            var path = UserOptionsHelper.GetOption<string>(UserOptionProperties.PluginRegistrationToolPath);

            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show(Resource.MessageBox_SetPRTOptionsPath);
                return;
            }

            if (!path.EndsWith("exe", StringComparison.CurrentCultureIgnoreCase))
                path = Path.Combine(path, "PluginRegistration.exe");

            if (!File.Exists(path))
            {
                MessageBox.Show($"{Resource.MessageBox_PRTNotFound}: " + path);
                return;
            }

            StartPrt(path);
        }

        private static void StartPrt(string path)
        {
            try
            {
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorLaunchingPRT, ex);
                MessageBox.Show(Resource.ErrorMessage_ErrorLaunchingPRT);
            }
        }
    }
}