using D365DeveloperExtensions.Core.Enums;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace D365DeveloperExtensions.Core.Models
{
    public class WebResourceType
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public int Type { get; set; }
        public int CrmMinimumMajorVersion { get; set; }
        public int CrmMaximumMajorVersion { get; set; }
        public bool AllowCompare { get; set; }
    }

    public static class WebResourceTypes
    {
        public static List<WebResourceType> Types => new List<WebResourceType> {
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "HTML", DisplayName = "Webpage (HTML)", Type = 1, AllowCompare = true},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "CSS", DisplayName = "Style Sheet (CSS)", Type = 2, AllowCompare = true},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "JS", DisplayName = "Script (JScript)", Type = 3, AllowCompare = true},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "XML", DisplayName = "Data (XML)", Type = 4, AllowCompare = true},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "PNG", DisplayName = "PNG format", Type = 5, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "JPG", DisplayName = "JPG format", Type = 6, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "GIF", DisplayName = "GIF format", Type = 7, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "XAP", DisplayName = "Silverlight (XAP)", Type = 8, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "XSL", DisplayName = "Style Sheet (XSL)", Type = 9, AllowCompare = true},
            new WebResourceType{CrmMinimumMajorVersion = 5, CrmMaximumMajorVersion = 99, Name = "ICO", DisplayName = "ICO format", Type = 10, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 9, CrmMaximumMajorVersion = 99, Name = "SVG", DisplayName = "SVG format", Type = 11, AllowCompare = false},
            new WebResourceType{CrmMinimumMajorVersion = 9, CrmMaximumMajorVersion = 99, Name = "RESX", DisplayName = "RESX format", Type = 12, AllowCompare = true},
        };

        public static ObservableCollection<WebResourceType> GetTypes(int majorVersion, bool addEmpty)
        {
            var types = new ObservableCollection<WebResourceType>(Types.Where(t =>
                t.CrmMinimumMajorVersion <= majorVersion && t.CrmMaximumMajorVersion >= majorVersion).ToList());

            if (addEmpty)
                types.Insert(0, new WebResourceType { CrmMinimumMajorVersion = 0, CrmMaximumMajorVersion = 99, Name = String.Empty, DisplayName = String.Empty, Type = -1 });

            return types;
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
                case "11":
                    return "SVG";
                case "12":
                    return "RESX";
                default:
                    return String.Empty;
            }
        }

        public static FileExtensionType GetExtensionType(string file)
        {
            string extension = Path.GetExtension(file);
            if (string.IsNullOrEmpty(extension))
                return FileExtensionType.None;

            extension = extension.Replace(".", String.Empty).ToUpper();

            switch (extension)
            {
                case "HTML":
                case "HTM":
                    return FileExtensionType.Html;
                case "CSS":
                    return FileExtensionType.Css;
                case "JS":
                    return FileExtensionType.Js;
                case "XML":
                    return FileExtensionType.Xml;
                case "PNG":
                    return FileExtensionType.Png;
                case "JPG":
                    return FileExtensionType.Jpg;
                case "GIF":
                    return FileExtensionType.Gif;
                case "XAP":
                    return FileExtensionType.Xap;
                case "XSL":
                    return FileExtensionType.Xsl;
                case "ICO":
                    return FileExtensionType.Ico;
                case "SVG":
                    return FileExtensionType.Svg;
                case "RESX":
                    return FileExtensionType.Resx;
                case "TS":
                    return FileExtensionType.Ts;
                case "MAP":
                    return FileExtensionType.Map;
                default:
                    return FileExtensionType.None;
            }
        }

        public static bool IsImageType(FileExtensionType extension)
        {
            List<FileExtensionType> imageExtensions = new List<FileExtensionType> {
                FileExtensionType.Ico,
                FileExtensionType.Png,
                FileExtensionType.Gif,
                FileExtensionType.Jpg,
                FileExtensionType.Svg };

            return imageExtensions.Any(i => i == extension);
        }
    }
}