using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Resources;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace D365DeveloperExtensions.Core.Crm
{
    public class Publish
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static bool PublishAllCustomizations(CrmServiceClient client)
        {
            try
            {
                ExLogger.LogToFile(Logger, Resource.Message_PublishingAllCustomizations, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_PublishingAllCustomizations, MessageType.Info);

                var request = new PublishAllXmlRequest();

                client.Execute(request);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorPublishing, ex);

                return false;
            }
        }
    }
}