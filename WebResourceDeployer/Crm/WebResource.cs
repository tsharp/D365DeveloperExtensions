using CrmDeveloperExtensions2.Core;
using CrmDeveloperExtensions2.Core.Enums;
using CrmDeveloperExtensions2.Core.Logging;
using CrmDeveloperExtensions2.Core.Models;
using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using WebResourceDeployer.Resources;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Crm
{
    public static class WebResource
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static EntityCollection RetrieveWebResourcesFromCrm(CrmServiceClient client)
        {
            EntityCollection results = null;
            try
            {
                int pageNumber = 1;
                string pagingCookie = null;
                bool moreRecords = true;

                while (moreRecords)
                {
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "solutioncomponent",
                        ColumnSet = new ColumnSet("solutionid"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "componenttype",
                                    Operator = ConditionOperator.Equal,
                                    Values = { 61 }
                                }
                            }
                        },
                        LinkEntities =
                        {
                            new LinkEntity
                            {
                                Columns = new ColumnSet("name", "displayname", "webresourcetype", "ismanaged", "webresourceid", "description"),
                                EntityAlias = "webresource",
                                LinkFromEntityName = "solutioncomponent",
                                LinkFromAttributeName = "objectid",
                                LinkToEntityName = "webresource",
                                LinkToAttributeName = "webresourceid"
                            }
                        },
                        PageInfo = new PagingInfo
                        {
                            PageNumber = pageNumber,
                            PagingCookie = pagingCookie
                        }
                    };

                    EntityCollection partialResults = client.RetrieveMultiple(query);

                    if (partialResults.MoreRecords)
                    {
                        pageNumber++;
                        pagingCookie = partialResults.PagingCookie;
                    }

                    moreRecords = partialResults.MoreRecords;

                    if (partialResults.Entities == null)
                        continue;

                    if (results == null)
                        results = new EntityCollection();

                    results.Entities.AddRange(partialResults.Entities);
                }

                OutputLogger.WriteToOutputWindow(Resource.Message_RetrievedWebResources, MessageType.Info);

                return results;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingWebResources, ex);

                return null;
            }
        }

        public static Entity RetrieveWebResourceFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                Entity webResource = client.Retrieve("webresource", webResourceId,
                    new ColumnSet("content", "name", "webresourcetype"));

                OutputLogger.WriteToOutputWindow($"{Resource.Message_DownloadedWebResource}: " + webResource.Id, MessageType.Info);

                return webResource;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingWebResource, ex);

                return null;
            }
        }

        public static Entity RetrieveWebResourceContentFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                Entity webResource = client.Retrieve("webresource", webResourceId, new ColumnSet("content", "name"));

                OutputLogger.WriteToOutputWindow($"{Resource.Message_RetrievedWebResourceContent}: {webResourceId}", MessageType.Info);

                return webResource;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingWebResourceContent, ex);

                return null;
            }
        }

        public static string RetrieveWebResourceDescriptionFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                Entity webResource = client.Retrieve("webresource", webResourceId, new ColumnSet("description"));

                return webResource.GetAttributeValue<string>("description");
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingWebResourceDescription, ex);

                return null;
            }
        }

        public static void DeleteWebResourcetFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                client.Delete("webresource", webResourceId);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorDeletingWebResource, ex);
            }
        }

        public static Entity CreateNewWebResourceEntity(int type, string prefix, string name, string displayName, string description, string filePath)
        {
            Entity webResource = new Entity("webresource")
            {
                ["name"] = prefix + name,
                ["webresourcetype"] = new OptionSetValue(type),
                ["description"] = description
            };

            if (!string.IsNullOrEmpty(displayName))
                webResource["displayname"] = displayName;

            if (type == 8)
                webResource["silverlightversion"] = "4.0";

            webResource["content"] = GetFileContent(filePath);

            return webResource;
        }

        private static string GetFileContent(string filePath)
        {
            FileExtensionType extension = WebResourceTypes.GetExtensionType(filePath);

            string content;

            //Images
            if (WebResourceTypes.IsImageType(extension))
            {
                content = EncodedImage(filePath, extension);
                return content;
            }

            //Everything else
            content = File.ReadAllText(filePath);

            return EncodeString(content);
        }

        public static Guid CreateWebResourceInCrm(CrmServiceClient client, Entity webResource)
        {
            try
            {
                Guid id = client.Create(webResource);

                OutputLogger.WriteToOutputWindow($"{Resource.Message_NewWebResourceCreated}: " + id, MessageType.Info);

                return id;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorCreatingWebResource, ex);

                return Guid.Empty;
            }
        }

        public static Entity CreateUpdateWebResourceEntity(Guid webResourceId, string boundFile, string description, string projectFullName)
        {
            Entity webResource = new Entity("webresource")
            {
                Id = webResourceId,
                ["description"] = description
            };

            string filePath = Path.GetDirectoryName(projectFullName) + boundFile.Replace("/", "\\");

            webResource["content"] = GetFileContent(filePath);

            return webResource;
        }

        public static bool UpdateAndPublishSingle(CrmServiceClient client, List<Entity> webResources)
        {
            //CRM 2011 < UR12
            try
            {
                OrganizationRequestCollection requests = CreateUpdateRequests(webResources);

                foreach (OrganizationRequest request in requests)
                {
                    client.Execute(request);
                    OutputLogger.WriteToOutputWindow(Resource.Message_UploadedWebResource, MessageType.Info);
                }

                string publishXml = CreatePublishXml(webResources);
                PublishXmlRequest publishRequest = CreatePublishRequest(publishXml);

                client.Execute(publishRequest);

                OutputLogger.WriteToOutputWindow(Resource.Message_PublishedWebResources, MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorUpdatingPublishingWebResources, ex);

                return false;
            }
        }

        public static bool UpdateAndPublishMultiple(CrmServiceClient client, List<Entity> webResources)
        {
            //CRM 2011 UR12+
            try
            {
                ExecuteMultipleRequest emRequest = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    }
                };

                emRequest.Requests = CreateUpdateRequests(webResources);

                string publishXml = CreatePublishXml(webResources);

                emRequest.Requests.Add(CreatePublishRequest(publishXml));

                bool wasError = false;
                ExecuteMultipleResponse emResponse = (ExecuteMultipleResponse)client.Execute(emRequest);

                foreach (var responseItem in emResponse.Responses)
                {
                    if (responseItem.Fault == null) continue;

                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_ErrorUpdatingPublishingWebResources, MessageType.Error);
                    wasError = true;
                }

                if (wasError)
                    return false;

                OutputLogger.WriteToOutputWindow(Resource.Message_UpdatedPublishedWebResources, MessageType.Info);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorUpdatingPublishingWebResources, ex);

                return false;
            }
        }

        private static PublishXmlRequest CreatePublishRequest(string publishXml)
        {
            PublishXmlRequest pubRequest = new PublishXmlRequest { ParameterXml = publishXml };
            return pubRequest;
        }

        private static OrganizationRequestCollection CreateUpdateRequests(List<Entity> webResources)
        {
            OrganizationRequestCollection requests = new OrganizationRequestCollection();

            foreach (Entity webResource in webResources)
            {
                UpdateRequest request = new UpdateRequest { Target = webResource };
                requests.Add(request);
            }

            return requests;
        }

        private static string CreatePublishXml(List<Entity> webResources)
        {
            StringBuilder publishXml = new StringBuilder();
            publishXml.Append("<importexportxml><webresources>");

            foreach (Entity webResource in webResources)
            {
                publishXml.Append($"<webresource>{webResource.Id}</webresource>");
            }

            publishXml.Append("</webresources></importexportxml>");

            return publishXml.ToString();
        }

        public static List<Entity> CreateDescriptionUpdateWebResource(WebResourceItem webResourceItem, string newDescription)
        {
            List<Entity> webResources = new List<Entity>();
            Entity webResource = new Entity("webresource")
            {
                Id = webResourceItem.WebResourceId,
                ["description"] = newDescription
            };

            webResources.Add(webResource);

            return webResources;
        }

        public static string GetWebResourceContent(Entity webResource)
        {
            bool hasContent = webResource.Attributes.TryGetValue("content", out var contentObj);
            var content = hasContent ? contentObj.ToString() : String.Empty;

            return content;
        }

        public static byte[] GetDecodedContent(Entity webResource)
        {
            string content = WebResource.GetWebResourceContent(webResource);
            byte[] decodedContent = WebResource.DecodeWebResource(content);

            return decodedContent;
        }

        public static byte[] DecodeWebResource(string value)
        {
            return Convert.FromBase64String(value);
        }

        public static string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        public static string EncodedImage(string filePath, FileExtensionType extension)
        {
            switch (extension)
            {
                case FileExtensionType.Ico:
                    return ImageEncoding.EncodeIco(filePath);
                case FileExtensionType.Gif:
                case FileExtensionType.Jpg:
                case FileExtensionType.Png:
                    return ImageEncoding.EncodeImage(filePath, extension);
                case FileExtensionType.Svg:
                    return ImageEncoding.EncodeSvg(filePath);
                default:
                    return null;
            }
        }

        public static string ConvertWebResourceNameToPath(string webResourceName, string folder, string projectFullName)
        {
            string[] name = webResourceName.Split('/');
            folder = folder.Replace("/", "\\");
            var path = Path.GetDirectoryName(projectFullName) +
                       (folder != "\\" ? folder : String.Empty) +
                       "\\" + name[name.Length - 1];

            return path;
        }

        [SuppressMessage("ReSharper", "AssignNullToNotNullAttribute")]
        public static string ConvertWebResourceNameFullToPath(string webResourceName, string rootFolder, Project project)
        {
            string[] folders = webResourceName.Split('/');

            string currentFullPath = Path.GetDirectoryName(project.FullName);
            string currentPartialPath = String.Empty;
            for (int i = 0; i < folders.Length - 1; i++)
            {
                string currentFolder = CrmDeveloperExtensions2.Core.Vs.ProjectItemWorker.CreateValidFolderName(folders[i]);
                currentFullPath = Path.Combine(currentFullPath, currentFolder);
                currentPartialPath = Path.Combine(currentPartialPath, currentFolder);
                bool exists = Directory.Exists(currentFullPath);
                if (!exists)
                    Directory.CreateDirectory(currentFullPath);

                CrmDeveloperExtensions2.Core.Vs.ProjectItemWorker.GetProjectItems(project, currentPartialPath, true);
            }

            return Path.Combine(currentFullPath, folders[folders.Length - 1]);
        }

        public static string AddMissingExtension(string name, int webResourceType)
        {
            if (!string.IsNullOrEmpty(Path.GetExtension(name)))
                return name;

            string ext = WebResourceTypes.GetWebResourceTypeNameByNumber(webResourceType.ToString()).ToLower();
            name += "." + ext;

            return name;
        }

        public static string GetExistingFolderFromBoundFile(WebResourceItem webResourceItem, string folder)
        {
            var directoryName = Path.GetDirectoryName(webResourceItem.BoundFile);
            if (directoryName != null)
                folder = directoryName.Replace("\\", "/");
            if (folder == "/")
                folder = String.Empty;
            return folder;
        }

        public static bool ValidateName(string name)
        {
            name = name.Trim();

            Regex r = new Regex("^[a-zA-Z0-9_.\\/]*$");
            if (!r.IsMatch(name))
                return false;

            return !name.Contains("//");
        }
    }
}