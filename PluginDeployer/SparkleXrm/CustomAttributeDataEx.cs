using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace SparkleXrm.Tasks
{
    public static class AttributeExtensions
    {
        public static CrmPluginRegistrationAttribute CreateFromData(this CustomAttributeData data)
        {
            CrmPluginRegistrationAttribute attribute = null;
            var arguments = data.ConstructorArguments.ToArray();
            // determine which constructor is being used by the first type
            if (data.ConstructorArguments.Count == 8 && data.ConstructorArguments[0].ArgumentType.Name == "String")
            {
                attribute = new CrmPluginRegistrationAttribute(
                    (string)arguments[0].Value,
                    (string)arguments[1].Value,
                    (StageEnum)Enum.ToObject(typeof(StageEnum), (int)arguments[2].Value),
                    (ExecutionModeEnum)Enum.ToObject(typeof(ExecutionModeEnum), (int)arguments[3].Value),
                    (string)arguments[4].Value,
                    (string)arguments[5].Value,
                    (int)arguments[6].Value,
                    (IsolationModeEnum)Enum.ToObject(typeof(IsolationModeEnum), (int)arguments[7].Value)
                    );

            }
            else if (data.ConstructorArguments.Count == 8 && data.ConstructorArguments[0].ArgumentType.Name == "MessageNameEnum")
            {
                attribute = new CrmPluginRegistrationAttribute(
                   (MessageNameEnum)Enum.ToObject(typeof(MessageNameEnum), (int)arguments[0].Value),
                   (string)arguments[1].Value,
                   (StageEnum)Enum.ToObject(typeof(StageEnum), (int)arguments[2].Value),
                   (ExecutionModeEnum)Enum.ToObject(typeof(ExecutionModeEnum), (int)arguments[3].Value),
                   (string)arguments[4].Value,
                   (string)arguments[5].Value,
                   (int)arguments[6].Value,
                   (IsolationModeEnum)Enum.ToObject(typeof(IsolationModeEnum), (int)arguments[7].Value)
                   );
            }
            else if (data.ConstructorArguments.Count == 5 && data.ConstructorArguments[0].ArgumentType.Name == "String")
            {
                attribute = new CrmPluginRegistrationAttribute(
                (string)arguments[0].Value,
                (string)arguments[1].Value,
                (string)arguments[2].Value,
                (string)arguments[3].Value,
                (IsolationModeEnum)Enum.ToObject(typeof(IsolationModeEnum), (int)arguments[4].Value)
                );
            }

            foreach (var namedArgument in data.NamedArguments)
            {
                switch (namedArgument.MemberName)
                {
                    case "Id":
                        attribute.Id = (string)namedArgument.TypedValue.Value;
                        break;
                    case "FriendlyName":
                        attribute.FriendlyName = (string)namedArgument.TypedValue.Value;
                        break;
                    case "GroupName":
                        attribute.FriendlyName = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Image1Name":
                        attribute.Image1Name = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Image1Attributes":
                        attribute.Image1Attributes = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Image2Name":
                        attribute.Image2Name = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Image2Attributes":
                        attribute.Image2Attributes = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Image1Type":
                        attribute.Image1Type = (ImageTypeEnum)namedArgument.TypedValue.Value;
                        break;
                    case "Image2Type":
                        attribute.Image2Type = (ImageTypeEnum)namedArgument.TypedValue.Value;
                        break;
                    case "Description":
                        attribute.Description = (string)namedArgument.TypedValue.Value;
                        break;
                    case "DeleteAsyncOperaton":
                        attribute.DeleteAsyncOperaton = (bool)namedArgument.TypedValue.Value;
                        break;
                    case "UnSecureConfiguration":
                        attribute.UnSecureConfiguration = (string)namedArgument.TypedValue.Value;
                        break;
                    case "SecureConfiguration":
                        attribute.SecureConfiguration = (string)namedArgument.TypedValue.Value;
                        break;
                    case "Offline":
                        attribute.Offline = (bool)namedArgument.TypedValue.Value;
                        break;
                    case "Server":
                        attribute.Server = (bool)namedArgument.TypedValue.Value;
                        break;
                    case "Action":
                        attribute.Action = (PluginStepOperationEnum)namedArgument.TypedValue.Value;
                        break;
                }
            }
            return attribute;
        }

        public static string GetAttributeCode(this CrmPluginRegistrationAttribute attribute, string indentation)
        {
            var code = string.Empty;
            var targetType = (attribute.Stage != null) ? TargetType.Plugin : TargetType.WorkflowAcitivty;

            string additionalParmeters = "";

            // Image 1
            if (attribute.Image1Name != null)
                additionalParmeters += ",Image1Type = ImageTypeEnum." + attribute.Image1Type;
            if (attribute.Image1Name != null)
                additionalParmeters += ",Image1Name = \"" + attribute.Image1Name + "\"";
            if (attribute.Image1Name != null)
                additionalParmeters += ",Image1Attributes = \"" + attribute.Image1Attributes + "\"";

            // Image 2
            if (attribute.Image2Name != null)
                additionalParmeters += ",Image2Type = ImageTypeEnum." + attribute.Image2Type;
            if (attribute.Image2Name != null)
                additionalParmeters += ",Image2Name = \"" + attribute.Image2Name + "\"";
            if (attribute.Image2Attributes != null)
                additionalParmeters += ",Image2Attributes = \"" + attribute.Image2Attributes + "\"";


            if (targetType == TargetType.Plugin)
            {
                // Description is only option for plugins
                if (attribute.Description != null)
                    additionalParmeters += ",Description = \"" + attribute.Description + "\"";
                if (attribute.Offline)
                    additionalParmeters += ",Offline = " + attribute.Offline;
                if (!attribute.Server)
                    additionalParmeters += ",Server = " + attribute.Server;
            }
            if (attribute.Id != null)
                additionalParmeters += ",Id = \"" + attribute.Id + "\"";

            if (attribute.DeleteAsyncOperaton != null)
                additionalParmeters += ",DeleteAsyncOperaton = " + attribute.DeleteAsyncOperaton;

            if (attribute.UnSecureConfiguration != null)
                additionalParmeters += ",UnSecureConfiguration = \"" + attribute.UnSecureConfiguration + "\"";

            if (attribute.SecureConfiguration != null)
                additionalParmeters += ",SecureConfiguration = \"" + attribute.SecureConfiguration + "\"";

            if (attribute.Action != null)
                additionalParmeters += ",Action = PluginStepOperationEnum." + attribute.Action;

            string tab = "    ";
            Regex parser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            // determine which template to use
            if (targetType == TargetType.Plugin)
            {
                // Plugin Step
                string template = "\"{0}\",\"{1}\",StageEnum.{2},ExecutionModeEnum.{3},\"{4}\",\"{5}\",{6},IsolationModeEnum.{7}{8}";

                code = String.Format(template,
                    attribute.Message,
                    attribute.EntityLogicalName,
                    attribute.Stage,
                    attribute.ExecutionMode,
                    attribute.FilteringAttributes,
                    attribute.Name,
                    attribute.ExecutionOrder,
                    attribute.IsolationMode,
                    additionalParmeters);
            }
            else
            {
                // Workflow Step
                string template = "\"{0}\",\"{1}\",\"{2}\",\"{3}\",IsolationModeEnum.{4}{5}";

                code = String.Format(template,
                    attribute.Name,
                    attribute.FriendlyName,
                    attribute.Description,
                    attribute.GroupName,
                    attribute.IsolationMode,
                    additionalParmeters);
            }

            String[] fields = parser.Split(code);
            code = String.Join($",{indentation}{tab}", fields);
            string regionName = targetType == TargetType.Plugin
                ? $"{attribute.Message} {attribute.EntityLogicalName}"
                : $"{attribute.Name}";
            code = $"{indentation}#region {attribute.Message}{regionName}{indentation}[CrmPluginRegistration({indentation}{tab}" + code + $"{indentation})]{indentation}#endregion";

            return code;
        }
    }
}