using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using SolutionPackager.Resources;
using SolutionPackager.ViewModels;
using System;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace SolutionPackager.Crm
{
    public static class Solution
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static EntityCollection RetrieveSolutionsFromCrm(CrmServiceClient client)
        {
            try
            {
                QueryExpression query = new QueryExpression
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
                            },
                            new ConditionExpression
                            {
                                AttributeName = "ismanaged",
                                Operator = ConditionOperator.Equal,
                                Values = { false }
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
                    Orders =
                    {
                        new OrderExpression
                        {
                            AttributeName = "friendlyname",
                            OrderType = OrderType.Ascending
                        }
                    }
                };

                EntityCollection solutions = client.RetrieveMultiple(query);

                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedSolutions, MessageType.Info);

                return solutions;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingSolutions, ex);

                return null;
            }
        }

        public static async Task<string> GetSolutionFromCrm(CrmServiceClient client, CrmSolution selectedSolution, bool managed)
        {
            try
            {
                // Hardcode connection timeout to one-hour to support large solutions.
                //TODO: Find a better way to handle this
                if (client.OrganizationServiceProxy != null)
                    client.OrganizationServiceProxy.Timeout = new TimeSpan(1, 0, 0);
                if (client.OrganizationWebProxyClient != null)
                    client.OrganizationWebProxyClient.InnerChannel.OperationTimeout = new TimeSpan(1, 0, 0);

                ExportSolutionRequest request = new ExportSolutionRequest
                {
                    Managed = managed,
                    SolutionName = selectedSolution.UniqueName
                };
                ExportSolutionResponse response = await Task.Run(() => (ExportSolutionResponse)client.Execute(request));

                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedSolution, MessageType.Info);

                string fileName = FileHandler.FormatSolutionVersionString(selectedSolution.UniqueName, selectedSolution.Version, managed);
                string tempFile = FileHandler.WriteTempFile(fileName, response.ExportSolutionFile);

                return tempFile;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingSolution, ex);

                return null;
            }
        }

        public static bool ImportSolution(CrmServiceClient client, string path)
        {
            byte[] solutionBytes = FileSystem.GetFileBytes(path);
            if (solutionBytes == null)
                return false;

            try
            {
                ImportSolutionRequest request = new ImportSolutionRequest
                {
                    CustomizationFile = solutionBytes,
                    OverwriteUnmanagedCustomizations = true,
                    PublishWorkflows = true,
                    ImportJobId = Guid.NewGuid()
                };

                client.Execute(request);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorImportingSolution, ex);

                return false;
            }
        }
    }
}