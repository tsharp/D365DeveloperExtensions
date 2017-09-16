using PluginDeployer.SparkleXrm;
using System.IO;
using System.Linq;

namespace PluginDeployer
{
    public class SpklHelpers
    {
        public static bool RegistrationDetailsPresent(string assemblyPath, bool isWorkflow)
        {
            var assemblyBytes = File.ReadAllBytes(assemblyPath);

            AssemblyContainer container = AssemblyContainer.LoadAssembly(assemblyBytes, isWorkflow, true);

            return container.PluginDatas.First().CrmPluginRegistrationAttributes.Count > 0;
        }
    }
}