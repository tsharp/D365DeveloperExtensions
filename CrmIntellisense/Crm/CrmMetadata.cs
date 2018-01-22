using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using CrmDeveloperExtensions2.Core.UserOptions;
using CrmIntellisense.Models;
using CrmIntellisense.Resources;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;

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
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievingMetadata, MessageType.Info);

                RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest
                {
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
                };

                RetrieveAllEntitiesResponse metaDataResponse = (RetrieveAllEntitiesResponse)client.Execute(metaDataRequest);

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
            OutputLogger.WriteToOutputWindow(Resource.Message_ProcessingMetadata, MessageType.Info);

            var entities = metaDataResponse.EntityMetadata;

            Metadata = new List<CompletionValue>();

            string entityTriggerCharacter = UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseEntityTriggerCharacter);
            string entityFieldCharacter = UserOptionsHelper.GetOption<string>(UserOptionProperties.IntellisenseFieldTriggerCharacter);

            foreach (EntityMetadata entityMetadata in entities)
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
            return label.LocalizedLabels.FirstOrDefault(l => l.LanguageCode == 1033)?.Label;
        }
    }
}