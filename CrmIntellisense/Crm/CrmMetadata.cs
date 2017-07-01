using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmIntellisense.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;

namespace CrmIntellisense.Crm
{
    public static class CrmMetadata
    {
        public static List<CompletionValue> Metadata { get; set; }

        public static void GetMetadata(CrmServiceClient client)
        {
            try
            {
                using (client)
                {
                    RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
                    };

                    RetrieveAllEntitiesResponse metaDataResponse = (RetrieveAllEntitiesResponse)client.Execute(metaDataRequest);

                    ProcessMetadata(metaDataResponse);
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Metadata From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Metadata From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
            }
        }

        private static void ProcessMetadata(RetrieveAllEntitiesResponse metaDataResponse)
        {
            var entities = metaDataResponse.EntityMetadata;

            Metadata = new List<CompletionValue>();

            foreach (EntityMetadata entityMetadata in entities)
            {
                Metadata.Add(new CompletionValue("$" + entityMetadata.LogicalName, entityMetadata.LogicalName, "Entity name"));
                Metadata.Add(new CompletionValue("$" + entityMetadata.LogicalName + "_?field?",
                    "$" + entityMetadata.LogicalName, "Entity field"));

                foreach (AttributeMetadata attribute in entityMetadata.Attributes)
                {
                    if (attribute.AttributeType != null)
                        Metadata.Add(new CompletionValue("$" + entityMetadata.LogicalName + "_" + attribute.LogicalName,
                            attribute.LogicalName, attribute.LogicalName + ": " + attribute.AttributeType.Value));
                }
            }

            Metadata = Metadata.OrderBy(m => m.Name).ToList();
        }
    }
}