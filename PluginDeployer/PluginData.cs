using System;
using System.Collections.Generic;
using System.Reflection;
using PluginDeployer.Spkl;

namespace PluginDeployer
{
    [Serializable]
    public class PluginData
    {
        public string AssemblyFullName { get; set; }
        public AssemblyName AssemblyName { get; set; }
        public List<CrmPluginRegistrationAttribute> CrmPluginRegistrationAttributes { get; set; }
    }
}