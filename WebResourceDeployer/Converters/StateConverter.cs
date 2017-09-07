using System;
using System.Globalization;
using System.Windows.Data;

namespace WebResourceDeployer.Converters
{
    public class StateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool state = Boolean.Parse(value.ToString());

            return state ? "Managed" : "Unmanaged";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
