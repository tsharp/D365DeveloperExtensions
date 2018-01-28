using Microsoft.Xrm.Tooling.Connector;

namespace D365DeveloperExtensions.Core.Connection
{
    public class ConnectEventArgs
    {
        public CrmServiceClient ServiceClient{ get; set; }
    }
}