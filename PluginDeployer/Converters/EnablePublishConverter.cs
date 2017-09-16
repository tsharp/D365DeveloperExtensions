using PluginDeployer.ViewModels;
using System;
using System.Globalization;
using System.Windows.Data;

namespace PluginDeployer.Converters
{
    public class EnablePublishConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            //Not connected
            bool hasConnected = Boolean.TryParse(values[0]?.ToString(), out bool isConnected);
            if (!hasConnected)
                return false;
            if (!isConnected)
                return false;

            //CrmAssembly
            CrmAssembly crmAssembly = null;
            if (values[1] != null)
                crmAssembly = (CrmAssembly)values[1];

            //DeploymentType
            bool hasDeploymentType = Int32.TryParse(values[2]?.ToString(), out int deploymentType);
            if (!hasDeploymentType)
                return false;

            if (deploymentType == 1)
                return true;

            return crmAssembly != null && crmAssembly.AssemblyId != Guid.Empty;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
