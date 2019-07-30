using CrmIntellisense.Models;
using CrmIntellisense.Resources;
using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
using D365DeveloperExtensions.Core.UserOptions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace CrmIntellisense.Crm
{
    public static class CrmMetadata
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static List<CompletionValue> Metadata { get; set; }


        public static void GetMetadata(CrmServiceClient client)
        {
            try
            {
                ExLogger.LogToFile(Logger, Resource.Message_RetrievingMetadata, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievingMetadata, MessageType.Info);

                var metaDataRequest = new RetrieveAllEntitiesRequest
                {
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
                };

                var metaDataResponse = (RetrieveAllEntitiesResponse)client.Execute(metaDataRequest);

                ExLogger.LogToFile(Logger, Resource.Message_RetrievedMetadata, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedMetadata, MessageType.Info);

                ProcessMetadata(metaDataResponse);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.Message_ErrorRetrievingMetadata, ex);
            }
        }

        private static void ProcessMetadata(RetrieveAllEntitiesResponse metaDataResponse)
        {
            ExLogger.LogToFile(Logger, Resource.Message_ProcessingMetadata, LogLevel.Info);
            OutputLogger.WriteToOutputWindow(Resource.Message_ProcessingMetadata, MessageType.Info);

            var entities = metaDataResponse.EntityMetadata;

            Metadata = new List<CompletionValue>();

            var entityTriggerCharacter = UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseEntityTriggerCharacter);
            var entityFieldCharacter = UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseFieldTriggerCharacter);

            foreach (var entityMetadata in entities)
            {
                Metadata.Add(new CompletionValue($"{entityTriggerCharacter}{entityMetadata.LogicalName}", entityMetadata.LogicalName,
                    GetDisplayName(entityMetadata.DisplayName), MetadataType.Entity));
                Metadata.Add(new CompletionValue($"{entityTriggerCharacter}{entityMetadata.LogicalName}{entityFieldCharacter}?field?",
                    $"{entityTriggerCharacter}{entityMetadata.LogicalName}", GetDisplayName(entityMetadata.DisplayName), MetadataType.None));

                foreach (var attribute in entityMetadata.Attributes.Where(attribute =>
                    attribute.IsValidForCreate.GetValueOrDefault() || attribute.IsValidForUpdate.GetValueOrDefault() || attribute.IsValidForRead.GetValueOrDefault()))
                {
                    Metadata.Add(new CompletionValue($"{entityTriggerCharacter}{entityMetadata.LogicalName}_{attribute.LogicalName}",
                        attribute.LogicalName, $"{GetDisplayName(attribute.DisplayName)}: {attribute.AttributeType.GetValueOrDefault()}", MetadataType.Attribute));
                }
            }

            Metadata = Metadata.OrderBy(m => m.Name).ToList();
        }

        private static string GetDisplayName(Label label)
        {
            // TODO: Adjust for localization
            return label.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == 1033)?.Label;
        }
    }
}