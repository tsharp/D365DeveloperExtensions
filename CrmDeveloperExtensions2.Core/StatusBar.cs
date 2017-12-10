using EnvDTE;

namespace CrmDeveloperExtensions2.Core
{
    public class StatusBar
    {
        private static DTE _dte;

        public StatusBar(DTE dte)
        {
            _dte = dte;
        }

        public static void SetStatusBarValue(string text)
        {
            _dte.StatusBar.Text = text;
        }

        public static void SetStatusBarValue(string text, vsStatusAnimation animation)
        {
            _dte.StatusBar.Text = text;
            _dte.StatusBar.Animate(true, animation);
        }

        public static void ClearStatusBarValue()
        {
            _dte.StatusBar.Clear();
        }

        public static void ClearStatusBarValue(vsStatusAnimation animation)
        {
            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, animation);
        }
    }
}