using System;
using System.Globalization;
using System.Windows.Data;

namespace PluginDeployer.Converters
{
    public class EnableGetAttributesButtonConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            //Not connected
            var hasConnected = bool.TryParse(values[0]?.ToString(), out var isConnected);
            if (!hasConnected)
                return false;
            if (!isConnected)
                return false;

            //DeploymentType
            var hasDeploymentType = int.TryParse(values[1]?.ToString(), out var deploymentType);
            if (!hasDeploymentType)
                return false;

            return deploymentType != 0;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}