using CrmDeveloperExtensions2.Core.Enums;
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

            System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(filePath);

            using (MemoryStream ms = new MemoryStream())
            {
                icon?.Save(ms);
                byte[] imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }

            return encodedImage;
        }

        public static string EncodeImage(string filePath, FileExtensionType extension)
        {
            string encodedImage;

            System.Drawing.Image image = System.Drawing.Image.FromFile(filePath, true);

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

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }

            return encodedImage;
        }

        public static string EncodeSvg(string filePath)
        {
            var inFile = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var binaryData = new Byte[inFile.Length];
            inFile.Read(binaryData, 0, (int)inFile.Length);
            inFile.Close();

            return Convert.ToBase64String(binaryData, 0, binaryData.Length);
        }
    }
}