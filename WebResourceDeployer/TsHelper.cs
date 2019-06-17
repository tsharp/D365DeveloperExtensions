using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Logging;
using EnvDTE;
using Newtonsoft.Json;
using NLog;
using System;
using System.IO;
using WebResourceDeployer.Models;
using WebResourceDeployer.Resources;

namespace WebResourceDeployer
{
    public static class TsHelper
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static string GetJsForTsPath(string filePath, Project project)
        {
            if (string.IsNullOrEmpty(filePath))
                return null;

            var jsPath = Path.ChangeExtension(filePath, "js");
            if (File.Exists(jsPath))
                return jsPath;

            var fileName = Path.GetFileName(jsPath);

            var tsConfigBytes = GetTsConfig(project);
            if (string.IsNullOrEmpty(tsConfigBytes))
                return null;

            var outDir = GetTsConfigOutDir(tsConfigBytes);

            return Path.Combine(outDir, fileName).Replace("\\", "/");
        }

        private static string GetTsConfigOutDir(string tsConfigBytes)
        {
            try
            {
                var tsConfig = JsonConvert.DeserializeObject<TsConfig>(tsConfigBytes);
                if (tsConfig.GetType().GetProperty("compilerOptions") == null)
                    return "/";
                var compilerOptions = tsConfig.compilerOptions;
                return compilerOptions.GetType().GetProperty("outDir") != null
                    ? !compilerOptions.outDir.StartsWith("/")
                        ? compilerOptions.outDir = "/" + compilerOptions.outDir
                        : compilerOptions.outDir
                    : "/";
            }
            catch (Exception ex)
            {
                ExceptionHandler.LogException(Logger, Resource.Error_UnableToGetTsconfigOutdir, ex);
                OutputLogger.WriteToOutputWindow(Resource.Error_UnableToGetTsconfigOutdir, MessageType.Error);
                return "/";
            }
        }

        private static string GetTsConfig(Project project)
        {
            var tsConfigPath = GetTsConfigPath(project);
            return File.Exists(tsConfigPath)
                ? File.ReadAllText(tsConfigPath)
                : null;
        }

        public static bool HasTsConfig(Project project)
        {
            var tsConfigPath = GetTsConfigPath(project);
            return File.Exists(tsConfigPath);
        }

        private static string GetTsConfigPath(Project project)
        {
            var projectPath = D365DeveloperExtensions.Core.Vs.ProjectWorker.GetProjectPath(project);
            return Path.Combine(projectPath, "tsconfig.json");
        }
    }
}