using System;
using System.ServiceModel;
using CrmDeveloperExtensions.Core.Enums;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace WebResourceDeployer.Crm
{
    public static class Solution
    {
        public static EntityCollection RetrieveSolutionsFromCrm(CrmServiceClient client, bool getManaged)
        {
            try
            {
                QueryExpression query = new QueryExpression
                {
                    EntityName = "solution",
                    ColumnSet = new ColumnSet("friendlyname", "solutionid", "uniquename"),
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
                    ConditionExpression noManaged = new ConditionExpression
                    {
                        AttributeName = "ismanaged",
                        Operator = ConditionOperator.Equal,
                        Values = { false }
                    };

                    query.Criteria.Conditions.Add(noManaged);
                }

                EntityCollection solutions = client.RetrieveMultiple(query);

                return solutions;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Solutions From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                CrmDeveloperExtensions.Core.Logging.OutputLogger.WriteToOutputWindow(
                    "Error Retrieving Solutions From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }
    }
}
