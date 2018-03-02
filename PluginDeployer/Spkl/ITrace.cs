using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginDeployer.Spkl
{
    public interface ITrace
    {
        void WriteLine(string format, params object[] args);
        void Write(string format, params object[] args);
    }
}
