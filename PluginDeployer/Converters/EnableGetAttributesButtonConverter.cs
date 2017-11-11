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
            bool hasConnected = Boolean.TryParse(values[0]?.ToString(), out bool isConnected);
            if (!hasConnected)
                return false;
            if (!isConnected)
                return false;

            //DeploymentType
            bool hasDeploymentType = Int32.TryParse(values[1]?.ToString(), out int deploymentType);
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