using D365DeveloperExtensions.Core.Enums;
using System;
using System.Drawing.Imaging;
using System.IO;

namespace WebResourceDeployer
{
    public class ImageEncoding
    {
        public static string EncodeIco(string filePath)
        {
            string encodedImage;

            var icon = System.Drawing.Icon.ExtractAssociatedIcon(filePath);

            using (var ms = new MemoryStream())
            {
                icon?.Save(ms);
                var imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }

            return encodedImage;
        }

        public static string EncodeImage(string filePath, FileExtensionType extension)
        {
            string encodedImage;

            var image = System.Drawing.Image.FromFile(filePath, true);

            ImageFormat format = null;
            switch (extension)
            {
                case FileExtensionType.Gif:
                    format = ImageFormat.Gif;
                    break;
                case FileExtensionType.Jpg:
                    format = ImageFormat.Jpeg;
                    break;
                case FileExtensionType.Png:
                    format = ImageFormat.Png;
                    break;
            }

            if (format == null)
                return null;

            using (var ms = new MemoryStream())
            {
                image.Save(ms, format);
                var imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }

            return encodedImage;
        }

        public static string EncodeSvg(string filePath)
        {
            var inFile = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var binaryData = new byte[inFile.Length];
            inFile.Read(binaryData, 0, (int)inFile.Length);
            inFile.Close();

            return Convert.ToBase64String(binaryData, 0, binaryData.Length);
        }
    }
}