using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

namespace $rootnamespace$
{
    public class $safeitemname$ : PluginBase
    {
        #region Constructor/Configuration
        private string _secureConfig = null;
        private string _unsecureConfig = null;

        public $safeitemname$(string unsecure, string secureConfig)
            : base(typeof($safeitemname$))
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecure;
        }
        #endregion

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            // TODO: Implement your custom code

        }
    }
}