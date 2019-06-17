using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using D365DeveloperExtensions.Core.Models;
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
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

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
                var pageNumber = 1;
                string pagingCookie = null;
                var moreRecords = true;

                while (moreRecords)
                {
                    var query = new QueryExpression
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

                    var partialResults = client.RetrieveMultiple(query);

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

                ExLogger.LogToFile(Logger, Resource.Message_RetrievedWebResources, LogLevel.Info);
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
                var webResource = client.Retrieve("webresource", webResourceId,
                    new ColumnSet("content", "name", "webresourcetype"));

                ExLogger.LogToFile(Logger, $"{Resource.Message_DownloadedWebResource}: {webResource.Id}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_DownloadedWebResource}: {webResource.Id}", MessageType.Info);

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
                var webResource = client.Retrieve("webresource", webResourceId, new ColumnSet("content", "name"));

                ExLogger.LogToFile(Logger, $"{Resource.Message_RetrievedWebResourceContent}: {webResourceId}", LogLevel.Info);
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
                var webResource = client.Retrieve("webresource", webResourceId, new ColumnSet("description"));

                ExLogger.LogToFile(Logger, $"{Resource.Message_RetrievedWebResourceDescription}: {webResourceId}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_RetrievedWebResourceDescription}: {webResourceId}", MessageType.Info);

                return webResource.GetAttributeValue<string>("description");
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorRetrievingWebResourceDescription, ex);

                return null;
            }
        }

        public static void DeleteWebResourceFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                client.Delete("webresource", webResourceId);

                ExLogger.LogToFile(Logger, $"{Resource.Message_DeletedWebResource}: {webResourceId}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_DeletedWebResource}: {webResourceId}", MessageType.Info);
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorDeletingWebResource, ex);
            }
        }

        public static Entity CreateNewWebResourceEntity(int type, string prefix, string name, string displayName, string description,
            string filePath, Project project)
        {
            var webResource = new Entity("webresource")
            {
                ["name"] = prefix + name,
                ["webresourcetype"] = new OptionSetValue(type),
                ["description"] = description
            };

            if (!string.IsNullOrEmpty(displayName))
                webResource["displayname"] = displayName;

            if (type == 8)
                webResource["silverlightversion"] = "4.0";

            webResource["content"] = GetFileContent(filePath, project);

            return webResource;
        }

        private static string GetFileContent(string filePath, Project project)
        {
            var extension = WebResourceTypes.GetExtensionType(filePath);

            //Images
            if (WebResourceTypes.IsImageType(extension))
            {
                var content = EncodedImage(filePath, extension);
                return content;
            }

            //TypeScript
            if (extension == FileExtensionType.Ts)
            {
                var jsPath = TsHelper.GetJsForTsPath(filePath, project);
                jsPath = FileSystem.BoundFileToLocalPath(jsPath,
                    D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project));
                return GetNonImageFileContext(jsPath);
            }

            //Everything else    
            return GetNonImageFileContext(filePath);
        }

        private static string GetNonImageFileContext(string filePath)
        {
            var content = FileSystem.GetFileText(filePath);
            return content == null
                ? null
                : EncodeString(content);
        }

        public static Guid CreateWebResourceInCrm(CrmServiceClient client, Entity webResource)
        {
            try
            {
                var id = client.Create(webResource);

                ExLogger.LogToFile(Logger, $"{Resource.Message_NewWebResourceCreated}: {id}", LogLevel.Info);
                OutputLogger.WriteToOutputWindow($"{Resource.Message_NewWebResourceCreated}: {id}", MessageType.Info);

                return id;
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.ErrorMessage_ErrorCreatingWebResource, ex);

                return Guid.Empty;
            }
        }

        public static Entity CreateUpdateWebResourceEntity(Guid webResourceId, string boundFile, string description, Project project)
        {
            var webResource = new Entity("webresource")
            {
                Id = webResourceId,
                ["description"] = description
            };

            var filePath = Path.GetDirectoryName(project.FullName) + boundFile.Replace("/", "\\");

            webResource["content"] = GetFileContent(filePath, project);

            return webResource;
        }

        public static bool UpdateAndPublishSingle(CrmServiceClient client, List<Entity> webResources)
        {
            //CRM 2011 < UR12
            try
            {
                var requests = CreateUpdateRequests(webResources);

                foreach (var request in requests)
                {
                    client.Execute(request);
                    OutputLogger.WriteToOutputWindow(Resource.Message_UploadedWebResource, MessageType.Info);
                }

                var publishXml = CreatePublishXml(webResources);
                var publishRequest = CreatePublishRequest(publishXml);

                client.Execute(publishRequest);

                ExLogger.LogToFile(Logger, Resource.Message_PublishedWebResources, LogLevel.Info);
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
                var emRequest = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    }
                };

                emRequest.Requests = CreateUpdateRequests(webResources);

                var publishXml = CreatePublishXml(webResources);

                emRequest.Requests.Add(CreatePublishRequest(publishXml));

                var wasError = false;
                var emResponse = (ExecuteMultipleResponse)client.Execute(emRequest);

                foreach (var responseItem in emResponse.Responses)
                {
                    if (responseItem.Fault == null) continue;

                    ExLogger.LogToFile(Logger, Resource.ErrorMessage_ErrorUpdatingPublishingWebResources, LogLevel.Info);
                    OutputLogger.WriteToOutputWindow(Resource.ErrorMessage_ErrorUpdatingPublishingWebResources, MessageType.Error);
                    wasError = true;
                }

                if (wasError)
                    return false;

                ExLogger.LogToFile(Logger, Resource.Message_UpdatedPublishedWebResources, LogLevel.Info);
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
            var pubRequest = new PublishXmlRequest { ParameterXml = publishXml };
            return pubRequest;
        }

        private static OrganizationRequestCollection CreateUpdateRequests(List<Entity> webResources)
        {
            var requests = new OrganizationRequestCollection();

            foreach (var webResource in webResources)
            {
                var request = new UpdateRequest { Target = webResource };
                requests.Add(request);
            }

            return requests;
        }

        private static string CreatePublishXml(List<Entity> webResources)
        {
            var publishXml = new StringBuilder();
            publishXml.Append("<importexportxml><webresources>");

            foreach (var webResource in webResources)
            {
                publishXml.Append($"<webresource>{webResource.Id}</webresource>");
            }

            publishXml.Append("</webresources></importexportxml>");

            return publishXml.ToString();
        }

        public static List<Entity> CreateDescriptionUpdateWebResource(WebResourceItem webResourceItem, string newDescription)
        {
            var webResources = new List<Entity>();
            var webResource = new Entity("webresource")
            {
                Id = webResourceItem.WebResourceId,
                ["description"] = newDescription
            };

            webResources.Add(webResource);

            return webResources;
        }

        public static string GetWebResourceContent(Entity webResource)
        {
            var hasContent = webResource.Attributes.TryGetValue("content", out var contentObj);
            var content = hasContent ? contentObj.ToString() : string.Empty;

            return content;
        }

        public static byte[] GetDecodedContent(Entity webResource)
        {
            var content = GetWebResourceContent(webResource);
            var decodedContent = DecodeWebResource(content);

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
            try
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
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow(Resource.Error_ErrorEncodingImage, MessageType.Error);
                ExceptionHandler.LogException(Logger, Resource.Error_ErrorEncodingImage, ex);
                return null;
            }
        }

        public static string ConvertWebResourceNameToPath(string webResourceName, string folder, string projectFullName)
        {
            var name = webResourceName.Split('/');
            folder = folder.Replace("/", "\\");
            var path = Path.GetDirectoryName(projectFullName) +
                       (folder != "\\" ? folder : string.Empty) +
                       "\\" + name[name.Length - 1];

            return path;
        }

        [SuppressMessage("ReSharper", "AssignNullToNotNullAttribute")]
        public static string ConvertWebResourceNameFullToPath(string webResourceName, string rootFolder, Project project)
        {
            var folders = webResourceName.Split('/');

            var currentFullPath = Path.GetDirectoryName(project.FullName);
            var currentPartialPath = string.Empty;
            for (var i = 0; i < folders.Length - 1; i++)
            {
                var currentFolder = D365DeveloperExtensions.Core.Vs.ProjectItemWorker.CreateValidFolderName(folders[i]);
                currentFullPath = Path.Combine(currentFullPath, currentFolder);
                currentPartialPath = Path.Combine(currentPartialPath, currentFolder);
                var exists = Directory.Exists(currentFullPath);
                if (!exists)
                    Directory.CreateDirectory(currentFullPath);

                D365DeveloperExtensions.Core.Vs.ProjectItemWorker.GetProjectItems(project, currentPartialPath, true);
            }

            return Path.Combine(currentFullPath, folders[folders.Length - 1]);
        }

        public static string AddMissingExtension(string name, int webResourceType)
        {
            if (!string.IsNullOrEmpty(Path.GetExtension(name)))
                return name;

            var ext = WebResourceTypes.GetWebResourceTypeNameByNumber(webResourceType.ToString()).ToLower();
            name += "." + ext;

            return name;
        }

        public static string GetExistingFolderFromBoundFile(WebResourceItem webResourceItem, string folder)
        {
            var directoryName = Path.GetDirectoryName(webResourceItem.BoundFile);
            if (directoryName != null)
                folder = directoryName.Replace("\\", "/");
            if (folder == "/")
                folder = string.Empty;
            return folder;
        }

        public static bool ValidateName(string name)
        {
            name = name.Trim();

            var r = new Regex("^[a-zA-Z0-9_.\\/]*$");
            if (!r.IsMatch(name))
                return false;

            return !name.Contains("//");
        }
    }
}