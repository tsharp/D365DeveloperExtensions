using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security;
using System.Security.Permissions;
using Microsoft.Xrm.Sdk;

namespace PluginDeployer.Spkl
{
    [Serializable]
    public class AssemblyContainer
    {
        public byte[] AssemblyData { get; set; }
        public bool ReflectionOnly { get; set; }
        private AppDomain Container { get; set; }
        private bool IsWorkflow { get; set; }
        public List<PluginData> PluginDatas { get; set; }


        public void Unload()
        {
            AppDomain.Unload(Container);
        }

        public void LoadAssembly()
        {
            PluginDatas = new List<PluginData>();

            var assembly = ReflectionOnly ? Assembly.ReflectionOnlyLoad(AssemblyData) : Assembly.Load(AssemblyData);

            IEnumerable<Type> pluginTypes;

            if (IsWorkflow)
                pluginTypes = Reflection.GetTypesInheritingFrom(assembly, typeof(System.Activities.CodeActivity));
            else
                pluginTypes = Reflection.GetTypesImplementingInterface(assembly, typeof(IPlugin));

            foreach (Type pluginType in pluginTypes)
            {
                PluginData pluginData = new PluginData { AssemblyName = assembly.GetName() };

                var attributes = Reflection.GetAttribute(pluginType, typeof(CrmPluginRegistrationAttribute).Name);
                pluginData.AssemblyFullName = pluginType.FullName;
                pluginData.CrmPluginRegistrationAttributes = new List<CrmPluginRegistrationAttribute>();
                foreach (CustomAttributeData data in attributes)
                {
                    pluginData.CrmPluginRegistrationAttributes.Add(data.CreateFromData());
                }

                PluginDatas.Add(pluginData);
            }

            Container.SetData("PluginDatas", PluginDatas);
        }

        /// <summary>
        /// Load the assembly into another domain
        /// </summary>
        /// <param name="assemblyBytes"></param>
        /// <param name="isWorkflow"></param>
        /// <param name="reflectionOnly"></param>
        /// <returns></returns>
        public static AssemblyContainer LoadAssembly(byte[] assemblyBytes, bool isWorkflow, bool reflectionOnly = false)
        {
            var containerAppDomain = AppDomain.CreateDomain(
                "AssemblyContainer",
                AppDomain.CurrentDomain.Evidence,
                new AppDomainSetup
                {
                    ApplicationBase = Environment.CurrentDirectory,
                    ShadowCopyFiles = "true"
                },
                new PermissionSet(PermissionState.Unrestricted));

            AssemblyContainer assemblyContainer = new AssemblyContainer()
            {
                AssemblyData = assemblyBytes,
                ReflectionOnly = reflectionOnly,
                Container = containerAppDomain,
                IsWorkflow = isWorkflow
            };

            containerAppDomain.DoCallBack(assemblyContainer.LoadAssembly);

            assemblyContainer.PluginDatas = containerAppDomain.GetData("PluginDatas") as List<PluginData>;

            return assemblyContainer;
        }
    }
}