using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;

namespace CrmDeveloperExtensions.Core.Connection
{
    public class ConnectEventArgs
    {
        public CrmServiceClient ServiceClient{ get; set; }
    }
}
