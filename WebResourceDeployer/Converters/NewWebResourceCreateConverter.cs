using System;
using System.Globalization;
using System.Windows.Data;

namespace WebResourceDeployer.Converters
{
    public class NewWebResourceCreateConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            for (var i = 0; i < values.Length; i++)
            {
                object o = values[i];
                bool isInt = Int32.TryParse(o.ToString(), out int v);
                if (isInt)
                {
                    if ((i == 0 || i == 2) && v == -1)
                        return false;
                    if (i == 1 && (v == -1 || v == 0))
                        return false;
                }

                if (string.IsNullOrEmpty(o.ToString()))
                    return false;
            }

            return true;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}