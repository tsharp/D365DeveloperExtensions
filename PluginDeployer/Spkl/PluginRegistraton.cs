using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using PluginDeployer.Spkl.Tasks;

namespace PluginDeployer.Spkl
{
    public class PluginRegistraton
    {
        private OrganizationServiceContext _ctx;
        private IOrganizationService _service;
        private ITrace _trace;
        private string[] _ignoredAssemblies = new string[] {
            "Microsoft.Crm.Sdk.Proxy.dll",
            "Microsoft.IdentityModel.dll",
            "Microsoft.Xrm.Sdk.dll",
            "Microsoft.Xrm.Sdk.Workflow.dll"
        };

        public PluginRegistraton(IOrganizationService service, OrganizationServiceContext context, ITrace trace)
        {
            _ctx = context;
            _service = service;
            _trace = trace;

        }
        /// <summary>
        /// If not null, components are added to this solution
        /// </summary>
        public string SolutionUniqueName { get; set; }

        public Guid RegisterWorkflowActivities(string file, string solutionName)
        {
            SolutionUniqueName = solutionName;
            var assemblyFilePath = new FileInfo(file);

            if (_ignoredAssemblies.Contains(assemblyFilePath.Name))
                return Guid.Empty;

            AssemblyContainer assemblyContainer = null;

            try
            {
                //Load the assembly in its own AppDomain to prevent load errors & file locking
                var assemblyBytes = File.ReadAllBytes(file);
                assemblyContainer = AssemblyContainer.LoadAssembly(assemblyBytes, true, true);
                List<PluginData> pluginDatas = assemblyContainer.PluginDatas;

                if (pluginDatas.Count <= 0)
                    return Guid.Empty;

                var plugin = RegisterAssembly(assemblyFilePath, pluginDatas.First().AssemblyName, pluginDatas.First().CrmPluginRegistrationAttributes);

                if (plugin == null)
                    return plugin.Id;

                foreach (PluginData pluginData in pluginDatas)
                {
                    RegisterActivities(pluginData.CrmPluginRegistrationAttributes, plugin, pluginData.AssemblyFullName);
                }

                _ctx.SaveChanges();

                return plugin.Id;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        private void RegisterActivities(List<CrmPluginRegistrationAttribute> crmPluginRegistrationAttributes, PluginAssembly plugin, string assemblyFullName)
        {
            var sdkPluginTypes = _ctx.GetPluginTypes(plugin);

            // Search for the CrmPluginStepAttribute
            if (!crmPluginRegistrationAttributes.Any())
                return;

            if (crmPluginRegistrationAttributes.Count > 1)
            {
                Debug.WriteLine("Workflow Activities can only have a single registration");
            }

            var workflowActivitiy = crmPluginRegistrationAttributes.First();

            // Check if the type is registered
            var sdkPluginType = sdkPluginTypes.FirstOrDefault(t => t.TypeName == assemblyFullName);

            if (sdkPluginType == null)
                sdkPluginType = new PluginType();

            // Update values
            sdkPluginType.Name = assemblyFullName;
            sdkPluginType.PluginAssemblyId = plugin.ToEntityReference();
            sdkPluginType.TypeName = assemblyFullName;
            sdkPluginType.FriendlyName = !string.IsNullOrEmpty(workflowActivitiy.FriendlyName) ? workflowActivitiy.FriendlyName : Guid.NewGuid().ToString();
            sdkPluginType.WorkflowActivityGroupName = workflowActivitiy.GroupName;
            sdkPluginType.Description = workflowActivitiy.Description;

            if (sdkPluginType.Id == Guid.Empty)
            {
                _trace.WriteLine("Registering Workflow Activity Type {0}", workflowActivitiy.Name);
                sdkPluginType.Id = _service.Create(sdkPluginType);
            }
            else
            {
                _trace.WriteLine("Updating Workflow Activity Type {0}", workflowActivitiy.Name);
                _ctx.UpdateObject(sdkPluginType);
            }
        }

        private void AddAssemblyToSolution(string solutionName, PluginAssembly assembly)
        {
            // Find solution
            AddSolutionComponentRequest addToSolution = new AddSolutionComponentRequest()
            {
                AddRequiredComponents = true,
                ComponentType = (int)componenttype.PluginAssembly,
                ComponentId = assembly.Id,
                SolutionUniqueName = solutionName
            };
            _trace.WriteLine("Adding to solution '{0}'", solutionName);
            _service.Execute(addToSolution);

        }

        private void AddTypeToSolution(string solutionName, PluginType sdkPluginType)
        {
            // Find solution
            AddSolutionComponentRequest addToSolution = new AddSolutionComponentRequest()
            {
                ComponentType = (int)componenttype.PluginType,
                ComponentId = sdkPluginType.Id,
                SolutionUniqueName = solutionName
            };
            _trace.WriteLine("Adding to solution '{0}'", solutionName);
            _service.Execute(addToSolution);
        }
        private void AddStepToSolution(string solutionName, SdkMessageProcessingStep sdkPluginType)
        {
            // Find solution
            AddSolutionComponentRequest addToSolution = new AddSolutionComponentRequest()
            {
                AddRequiredComponents = false,
                ComponentType = (int)componenttype.SDKMessageProcessingStep,
                ComponentId = sdkPluginType.Id,
                SolutionUniqueName = solutionName
            };
            _trace.WriteLine("Adding to solution '{0}'", solutionName);
            _service.Execute(addToSolution);

        }

        public Guid RegisterPlugin(string file, string solutionName)
        {
            SolutionUniqueName = solutionName;
            var assemblyFilePath = new FileInfo(file);
            AssemblyContainer assemblyContainer = null;

            try
            {
                //Load the assembly in its own AppDomain to prevent load errors & file locking
                var assemblyBytes = File.ReadAllBytes(file);
                assemblyContainer = AssemblyContainer.LoadAssembly(assemblyBytes, false, true);
                List<PluginData> pluginDatas = assemblyContainer.PluginDatas;

                if (pluginDatas.Count <= 0)
                    return Guid.Empty;

                var plugin = RegisterAssembly(assemblyFilePath, pluginDatas.First().AssemblyName, pluginDatas.First().CrmPluginRegistrationAttributes);

                if (plugin == null)
                    return Guid.Empty;

                foreach (PluginData pluginData in pluginDatas)
                {
                    RegisterPluginSteps(pluginData.CrmPluginRegistrationAttributes, plugin, pluginData.AssemblyFullName);
                }

                _ctx.SaveChanges();

                return plugin.Id;
            }
            finally
            {
                assemblyContainer?.Unload();
            }
        }

        private PluginAssembly RegisterAssembly(FileInfo assemblyFilePath, AssemblyName assembly, List<CrmPluginRegistrationAttribute> crmPluginRegistrationAttributes)
        {
            // Get the isolation mode of the first attribute
            var firstTypeAttribute = crmPluginRegistrationAttributes.First();

            // Is there any steps to register?
            if (firstTypeAttribute == null)
                return null;
            var assemblyProperties = assembly.FullName.Split(",= ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var assemblyName = assembly.Name;

            // If found then register or update it
            var plugin = (from p in _ctx.CreateQuery<PluginAssembly>()
                          where p.Name == assemblyName
                          select new PluginAssembly
                          {
                              Id = p.Id,
                              Name = p.Name
                          }).FirstOrDefault();

            string assemblyBase64 = Convert.ToBase64String(File.ReadAllBytes(assemblyFilePath.FullName));

            if (plugin == null)
            {
                plugin = new PluginAssembly();
            }

            // update
            plugin.Content = assemblyBase64;
            plugin.Name = assemblyProperties[0];
            plugin.Culture = assemblyProperties[4];
            plugin.Version = assemblyProperties[2];
            plugin.PublicKeyToken = assemblyProperties[6];
            plugin.SourceType = new OptionSetValue(0); // database
            plugin.IsolationMode = firstTypeAttribute.IsolationMode == IsolationModeEnum.Sandbox ? new OptionSetValue(2) : new OptionSetValue(1); // 1= node, 2 = sandbox

            if (plugin.Id == Guid.Empty)
            {
                _trace.WriteLine("Registering Plugin '{0}' from '{1}'", plugin.Name, assemblyFilePath.FullName);
                // Create
                plugin.Id = _service.Create(plugin);
            }
            else
            {
                _trace.WriteLine("Updating Plugin '{0}' from '{1}'", plugin.Name, assemblyFilePath.FullName);
                // Update
                _ctx.UpdateObject(plugin);
            }

            // Add to solution
            if (SolutionUniqueName != null)
            {
                _trace.WriteLine("Adding Plugin '{0}' to solution '{1}'", plugin.Name, SolutionUniqueName);
                AddAssemblyToSolution(SolutionUniqueName, plugin);
            }

            return plugin;
        }

        private void RegisterPluginSteps(List<CrmPluginRegistrationAttribute> crmPluginRegistrationAttributes, PluginAssembly plugin, string assemblyFullName)
        {
            var sdkPluginTypes = _ctx.GetPluginTypes(plugin);

            // Check if the type is registered
            var sdkPluginType = sdkPluginTypes.FirstOrDefault(t => t.TypeName == assemblyFullName);

            if (sdkPluginType == null)
                sdkPluginType = new PluginType();

            // Update values
            sdkPluginType.Name = assemblyFullName;
            sdkPluginType.PluginAssemblyId = plugin.ToEntityReference();
            sdkPluginType.TypeName = assemblyFullName;
            sdkPluginType.FriendlyName = assemblyFullName;

            if (sdkPluginType.Id == Guid.Empty)
            {
                _trace.WriteLine("Registering Type '{0}'", sdkPluginType.Name);
                sdkPluginType.Id = _service.Create(sdkPluginType);
            }
            else
            {
                _trace.WriteLine("Updating Type '{0}'", sdkPluginType.Name);
                _ctx.UpdateObject(sdkPluginType);
            }

            var existingSteps = GetExistingSteps(sdkPluginType);

            foreach (var pluginAttribute in crmPluginRegistrationAttributes)
            {
                RegisterStep(sdkPluginType, existingSteps, pluginAttribute);
            }
        }

        private List<SdkMessageProcessingStep> GetExistingSteps(PluginType sdkPluginType)
        {
            // Get existing Steps
            var steps = (from s in _ctx.CreateQuery<SdkMessageProcessingStep>()
                         where s.PluginTypeId.Id == sdkPluginType.Id
                         select new SdkMessageProcessingStep()
                         {
                             Id = s.Id,
                             PluginTypeId = s.PluginTypeId,
                             SdkMessageId = s.SdkMessageId,
                             Mode = s.Mode,
                             Name = s.Name,
                             Rank = s.Rank,
                             Configuration = s.Configuration,
                             Description = s.Description,
                             Stage = s.Stage,
                             SupportedDeployment = s.SupportedDeployment,
                             FilteringAttributes = s.FilteringAttributes,
                             EventHandler = s.EventHandler,
                             AsyncAutoDelete = s.AsyncAutoDelete,
                             Attributes = s.Attributes,
                             SdkMessageFilterId = s.SdkMessageFilterId

                         }).ToList();

            return steps;

        }

        private void RegisterStep(PluginType sdkPluginType, IEnumerable<SdkMessageProcessingStep> existingSteps, CrmPluginRegistrationAttribute pluginStep)

        {
            SdkMessageProcessingStep step = null;
            if (pluginStep.Id != null)
            {
                Guid stepId = new Guid(pluginStep.Id);
                // Get by ID
                step = existingSteps.Where(s => s.Id == stepId).FirstOrDefault();
            }

            if (step == null)
            {
                // Get by Name
                step = existingSteps.Where(s => s.Name == pluginStep.Name && s.SdkMessageId.Name == pluginStep.Message).FirstOrDefault();
            }

            // Register images
            if (step == null)
            {
                step = new SdkMessageProcessingStep();
            }
            Guid? sdkMessageId = null;
            Guid? sdkMessagefilterId = null;

            if (pluginStep.EntityLogicalName == "none")
            {
                var message = _ctx.GetMessage(pluginStep.Message);
                sdkMessageId = message.SdkMessageId;
            }
            else
            {
                var messageFilter = _ctx.GetMessageFilter(pluginStep.EntityLogicalName, pluginStep.Message);

                if (messageFilter == null)
                {
                    _trace.WriteLine("Warning: Cannot register step {0} on Entity {1}", pluginStep.Message, pluginStep.EntityLogicalName);
                    return;
                }

                sdkMessageId = messageFilter.SdkMessageId.Id;
                sdkMessagefilterId = messageFilter.SdkMessageFilterId;
            }

            // Update attributes
            step.Name = pluginStep.Name;
            step.Configuration = pluginStep.UnSecureConfiguration;
            step.Description = pluginStep.Description;
            step.Mode = new OptionSetValue(pluginStep.ExecutionMode == ExecutionModeEnum.Asynchronous ? 1 : 0);
            step.Rank = pluginStep.ExecutionOrder;
            int stage = 10;
            switch (pluginStep.Stage)
            {
                case StageEnum.PreValidation:
                    stage = 10;
                    break;
                case StageEnum.PreOperation:
                    stage = 20;
                    break;
                case StageEnum.PostOperation:
                    stage = 40;
                    break;
            }

            step.Stage = new OptionSetValue(stage);
            int supportDeployment = 0;
            if (pluginStep.Server == true && pluginStep.Offline == true)
            {
                supportDeployment = 2; // Both
            }
            else if (!pluginStep.Server == true && pluginStep.Offline == true)
            {
                supportDeployment = 1; // Offline only
            }
            else
                supportDeployment = 0; // Server Only
            step.SupportedDeployment = new OptionSetValue(supportDeployment);
            step.PluginTypeId = sdkPluginType.ToEntityReference();
            step.SdkMessageFilterId = sdkMessagefilterId != null ? new EntityReference(SdkMessageFilter.EntityLogicalName, sdkMessagefilterId.Value) : null;
            step.SdkMessageId = new EntityReference(SdkMessage.EntityLogicalName, sdkMessageId.Value);
            step.FilteringAttributes = pluginStep.FilteringAttributes;
            if (step.Id == Guid.Empty)
            {
                _trace.WriteLine("Registering Step '{0}'", step.Name);
                // Create
                step.Id = _service.Create(step);
            }
            else
            {
                _trace.WriteLine("Updating Step '{0}'", step.Name);
                // Update
                _ctx.UpdateObject(step);
            }

            // Get existing Images
            SdkMessageProcessingStepImage[] existingImages = _ctx.GetPluginStepImages(step);

            var image1 = RegisterImage(pluginStep, step, existingImages, pluginStep.Image1Name, pluginStep.Image1Type, pluginStep.Image1Attributes);
            var image2 = RegisterImage(pluginStep, step, existingImages, pluginStep.Image2Name, pluginStep.Image2Type, pluginStep.Image2Attributes);

            if (SolutionUniqueName != null)
            {
                AddStepToSolution(SolutionUniqueName, step);

            }
        }




        private SdkMessageProcessingStepImage RegisterImage(CrmPluginRegistrationAttribute stepAttribute, SdkMessageProcessingStep step, SdkMessageProcessingStepImage[] existingImages, string imageName, ImageTypeEnum imagetype, string attributes)
        {
            if (String.IsNullOrWhiteSpace(imageName))
            {
                return null;
            }
            var image = existingImages.Where(a =>
                            a.SdkMessageProcessingStepId.Id == step.Id
                            &&
                            a.EntityAlias == imageName
                            && a.ImageType.Value == (int)imagetype).FirstOrDefault();
            if (image == null)
            {
                image = new SdkMessageProcessingStepImage();
            }

            image.Name = imageName;

            image.ImageType = new OptionSetValue((int)imagetype);
            image.SdkMessageProcessingStepId = new EntityReference(SdkMessageProcessingStep.EntityLogicalName, step.Id);
            image.Attributes1 = attributes;
            image.EntityAlias = imageName;
            image.MessagePropertyName = stepAttribute.Message == "Create" ? "Id" : "Target";
            if (image.Id == Guid.Empty)
            {
                _trace.WriteLine("Registering Image '{0}'", image.Name);
                image.Id = _service.Create(image);
            }
            else
            {
                _trace.WriteLine("Updating Image '{0}'", image.Name);
                _ctx.UpdateObject(image);
            }
            return image;
        }
    }
}
