using System.Threading;

namespace PluginDeployer.Spkl
{
    public class TraceLogger : ITrace
    {
        public void WriteLine(string format, params object[] args)
        {
            //if (format == null)
            //    return;

            //Console.WriteLine(format, args);
            //string logFile = GetLogFile();
            //using (var file = File.AppendText(GetLogFile()))
            //{
            //    file.WriteLine(DateTime.Now.ToLongTimeString() + "\t" + format, args);
            //    file.Flush();
            //}
        }

        public void Write(string format, params object[] args)
        {
            //Console.Write(format, args);
            //using (var file = File.AppendText(GetLogFile()))
            //{
            //    file.Write(DateTime.Now.ToLongTimeString() + "\t" + format, args);
            //    file.Flush();
            //}
        }
        private string GetLogFile()
        {
            return "Log" + Thread.CurrentThread.ManagedThreadId.ToString() + ".txt";
        }
    }
}
