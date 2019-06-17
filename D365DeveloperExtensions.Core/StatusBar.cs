using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NLog;
using System;

namespace D365DeveloperExtensions.Core
{
    public class StatusBar
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void SetStatusBarValue(string text)
        {
            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                dte.StatusBar.Text = text;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }

        public static void SetStatusBarValue(string text, vsStatusAnimation animation)
        {
            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                dte.StatusBar.Text = text;
                dte.StatusBar.Animate(true, animation);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }

        public static void ClearStatusBarValue()
        {
            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                dte.StatusBar.Clear();
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }

        public static void ClearStatusBarValue(vsStatusAnimation animation)
        {
            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                    throw new ArgumentNullException(Resources.Resource.ErrorMessage_ErrorAccessingDTE);

                dte.StatusBar.Clear();
                dte.StatusBar.Animate(false, animation);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resources.Resource.ErrorMessage_ErrorAccessingDTE, ex);
                throw;
            }
        }
    }
}