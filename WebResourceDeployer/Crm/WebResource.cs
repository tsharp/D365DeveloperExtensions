using System;
using System.Drawing.Imaging;
using System.IO;
using System.ServiceModel;
using System.Text;
using CrmDeveloperExtensions.Core.Enums;
using CrmDeveloperExtensions.Core.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using WebResourceDeployer.ViewModels;

namespace WebResourceDeployer.Crm
{
    public static class WebResource
    {
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
                                Columns = new ColumnSet("name", "displayname", "webresourcetype", "ismanaged", "webresourceid"),
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

                    if (partialResults.Entities == null) continue;

                    if (results == null)
                        results = new EntityCollection();

                    results.Entities.AddRange(partialResults.Entities);
                }

                return results;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resources From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resources From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }

        public static Entity RetrieveWebResourceFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                Entity webResource = client.Retrieve("webresource", webResourceId,
                    new ColumnSet("content", "name", "webresourcetype"));

                return webResource;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }

        public static Entity RetrieveWebResourceContentFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                Entity webResource = client.Retrieve("webresource", webResourceId, new ColumnSet("content", "name"));

                return webResource;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
                return null;
            }
        }

        public static void DeleteWebResourcetFromCrm(CrmServiceClient client, Guid webResourceId)
        {
            try
            {
                client.Delete("webresource", webResourceId);
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, MessageType.Error);
            }
            catch (Exception ex)
            {
                OutputLogger.WriteToOutputWindow("Error Retrieving Web Resource From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, MessageType.Error);
            }
        }

        public static string GetWebResourceTypeNameByNumber(string type)
        {
            switch (type)
            {
                case "1":
                    return "HTML";
                case "2":
                    return "CSS";
                case "3":
                    return "JS";
                case "4":
                    return "XML";
                case "5":
                    return "PNG";
                case "6":
                    return "JPG";
                case "7":
                    return "GIF";
                case "8":
                    return "XAP";
                case "9":
                    return "XSL";
                case "10":
                    return "ICO";
                default:
                    return String.Empty;
            }
        }

        public static string GetWebResourceContent(Entity webResource)
        {
            object contentObj;
            bool hasContent = webResource.Attributes.TryGetValue("content", out contentObj);
            var content = hasContent ? contentObj.ToString() : String.Empty;

            return content;
        }

        public static byte[] DecodeWebResource(string value)
        {
            return Convert.FromBase64String(value);
        }

        public static string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        public static string EncodedImage(string filePath, string extension)
        {
            string encodedImage;

            if (extension.ToUpper() == ".ICO")
            {
                System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(filePath);

                using (MemoryStream ms = new MemoryStream())
                {
                    icon?.Save(ms);
                    byte[] imageBytes = ms.ToArray();
                    encodedImage = Convert.ToBase64String(imageBytes);
                }

                return encodedImage;
            }

            System.Drawing.Image image = System.Drawing.Image.FromFile(filePath, true);

            ImageFormat format = null;
            switch (extension.ToUpper())
            {
                case ".GIF":
                    format = ImageFormat.Gif;
                    break;
                case ".JPG":
                    format = ImageFormat.Jpeg;
                    break;
                case ".PNG":
                    format = ImageFormat.Png;
                    break;
            }

            if (format == null)
                return null;

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }
            return encodedImage;
        }


        public static string ConvertWebResourceNameToPath(string webResourceName, string folder, string projectFullName)
        {
            string[] name = webResourceName.Split('/');
            folder = folder.Replace("/", "\\");
            var path = Path.GetDirectoryName(projectFullName) +
                       ((folder != "\\") ? folder : String.Empty) +
                       "\\" + name[name.Length - 1];

            return path;
        }

        public static string AddMissingExtension(string name, int webResourceType)
        {
            if (!string.IsNullOrEmpty(Path.GetExtension(name)))
                return name;

            string ext =
                GetWebResourceTypeNameByNumber(webResourceType.ToString()).ToLower();
            name += "." + ext;

            return name;
        }

        public static string GetExistingFolderFromBoundFile(WebResourceItem webResourceItem, string folder)
        {
            var directoryName = System.IO.Path.GetDirectoryName(webResourceItem.BoundFile);
            if (directoryName != null)
                folder = directoryName.Replace("\\", "/");
            if (folder == "/")
                folder = String.Empty;
            return folder;
        }
    }
}
