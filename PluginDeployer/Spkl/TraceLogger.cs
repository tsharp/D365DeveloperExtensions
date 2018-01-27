using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;

namespace PluginDeployer.Spkl
{
    public class TraceLogger : ITrace
    {
        public void WriteLine(string format, params object[] args)
        {
            if (format == null)
                return;

            OutputLogger.WriteToOutputWindow(string.Format(format, args), MessageType.Info);
        }

        public void Write(string format, params object[] args)
        {

        }
    }
}
