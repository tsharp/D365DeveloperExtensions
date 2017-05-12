using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace CrmDeveloperExtensions.Core
{
    public static class StatusBar
    {
        public static void SetStatusBarValue(DTE dte, string text)
        {
            dte.StatusBar.Text = text;
        }

        //public static void SetStatusBarValue(DTE dte, string text, int delayMs)
        //{
        //    dte.StatusBar.Text = text;
        //}

        public static void SetStatusBarValue(DTE dte, string text, vsStatusAnimation animation)
        {
            dte.StatusBar.Text = text;
            dte.StatusBar.Animate(true, animation);
        }

        public static void ClearStatusBarValue(DTE dte)
        {
            dte.StatusBar.Clear();
        }

        public static void ClearStatusBarValue(DTE dte, vsStatusAnimation animation)
        {
            dte.StatusBar.Clear();
            dte.StatusBar.Animate(false, animation);
        }
    }
}
