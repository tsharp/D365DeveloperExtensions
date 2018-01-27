using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginDeployer.Spkl.Config;

namespace PluginDeployer.Spkl.Config
{
    public class WebresourceDeployConfig
    {
        public string profile;
        public string root;
        public string solution;
        public List<WebResourceFile> files;
    }

}
