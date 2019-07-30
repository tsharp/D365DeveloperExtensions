using D365DeveloperExtensions.Core.Enums;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Runtime.InteropServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace D365DeveloperExtensions.Core
{
    public class MonikerHelper
    {
        public static ImageSource GetImage(MetadataType metadataType)
        {

            ImageMoniker imageMoniker;
            switch (metadataType)
            {
                case MetadataType.Attribute:
                    imageMoniker = KnownMonikers.Field;
                    break;
                case MetadataType.Entity:
                    imageMoniker = KnownMonikers.Table;
                    break;
                default:
                    return null;
            }

            return MonikerToBitmap(imageMoniker, 16);
        }

        private static BitmapSource MonikerToBitmap(ImageMoniker moniker, int size)
        {
            var shell = (IVsUIShell5)ServiceProvider.GlobalProvider.GetService(typeof(SVsUIShell));
            var backgroundColor = shell.GetThemedColorRgba(EnvironmentColors.MainWindowButtonActiveBorderBrushKey);

            var imageAttributes = new ImageAttributes
            {
                Flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags | unchecked((uint)_ImageAttributesFlags.IAF_Background),
                ImageType = (uint)_UIImageType.IT_Bitmap,
                Format = (uint)_UIDataFormat.DF_WPF,
                Dpi = 96,
                LogicalHeight = size,
                LogicalWidth = size,
                Background = backgroundColor,
                StructSize = Marshal.SizeOf(typeof(ImageAttributes))
            };

            var service = (IVsImageService2)Package.GetGlobalService(typeof(SVsImageService));
            var result = service.GetImage(moniker, imageAttributes);
            result.get_Data(out var data);

            return data as BitmapSource;
        }
    }
}