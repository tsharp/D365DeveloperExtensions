using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using WebResourceDeployer.Resources;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace WebResourceDeployer.Crm
{
    public static class Solution
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static EntityCollection RetrieveSolutionsFromCrm(CrmServiceClient client, bool getManaged)
        {
            try
            {
                var query = new QueryExpression
                {
                    EntityName = "solution",
                    ColumnSet = new ColumnSet("friendlyname", "solutionid", "uniquename", "version"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "isvisible",
                                Operator = ConditionOperator.Equal,
                                Values = {true}
                            }
                        }
                    },
                    LinkEntities =
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "solution",
                            LinkFromAttributeName = "publisherid",
                            LinkToEntityName = "publisher",
                            LinkToAttributeName = "publisherid",
                            Columns = new ColumnSet("customizationprefix"),
                            EntityAlias = "publisher"
                        }
                    },
                    Distinct = true,
                    Orders =
                    {
                        new OrderExpression
                        {
                            AttributeName = "friendlyname",
                            OrderType = OrderType.Ascending
                        }
                    }
                };

                if (!getManaged)
                {
                    var noManaged = new ConditionExpression
                    {
                        AttributeName = "ismanaged",
                        Operator = ConditionOperator.Equal,
                        Values = { false }
                    };

                    query.Criteria.Conditions.Add(noManaged);
                }

                var solutions = client.RetrieveMultiple(query);

                ExLogger.LogToFile(Logger, Resource.Message_RetrievedSolutions, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedSolutions, MessageType.Info);

                return solutions;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingSolutions, ex);

                return null;
            }
        }

        public static bool AddWebResourceToSolution(CrmServiceClient client, string uniqueName, Guid webResourceId)
        {
            try
            {
                var scRequest = new AddSolutionComponentRequest
                {
                    ComponentType = 61,
                    SolutionUniqueName = uniqueName,
                    ComponentId = webResourceId
                };

                client.Execute(scRequest);

                ExLogger.LogToFile(Logger, Resource.Message_NewWebResourceAddedSolution, LogLevel.Info);
                OutputLogger.WriteToOutputWindow(Resource.Message_NewWebResourceAddedSolution, MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorAddingWebResourceSolution, ex);

                return false;
            }
        }
    }
}