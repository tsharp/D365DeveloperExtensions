using CrmDeveloperExtensions2.Core.Models;
using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace WebResourceDeployer.Converters
{
    public class AllowCompareConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture) {
            //No bound file - nothing to compare to
            if (string.IsNullOrEmpty(values[0]?.ToString()))
                return false;

            bool hasType = Int32.TryParse(values[1].ToString(), out int type);
            if (!hasType)
                return false;

            WebResourceType webResourceType = WebResourceTypes.Types.FirstOrDefault(t => t.Type == type);

            return webResourceType != null && webResourceType.AllowCompare;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
