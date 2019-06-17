using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using NLog;
using NuGet;
using System;
using System.Collections.Generic;
using System.Linq;
using ExLogger = D365DeveloperExtensions.Core.Logging.ExtensionLogger;
using Logger = NLog.Logger;

namespace NuGetRetriever
{
    public static class PackageLister
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>  Retrieves NuGet package info.</summary>
        /// <param name="packageId">The NuGet package identifier.</param>
        /// <returns><![CDATA[ List<NuGetPackage> ]]>.</returns>
        public static List<NuGetPackage> GetPackagesById(string packageId)
        {
            try
            {
                ExLogger.LogToFile(Logger, $"{Resources.Resource.Message_RetrievingNuGetpackage}: {packageId}", LogLevel.Info);

                var packages = GetPackages(packageId);

                var results = new List<NuGetPackage>();
                foreach (var package in packages)
                {
                    if (package.Published != null && package.Published.Value.Year == 1900)
                        continue;

                    results.Add(CreateNuGetPackage(package));
                }

                ExLogger.LogToFile(Logger, $"Found {results.Count} packages", LogLevel.Info);

                return new List<NuGetPackage>(results.OrderByDescending(v => v.Version));
            }
            catch (Exception e)
            {
                ExceptionHandler.LogException(Logger, $"{Resources.Resource.ErrorMessage_FailedretrievingNuGetpackage}: {packageId}", e);
                throw;
            }
        }

        private static List<IPackage> GetPackages(string packageId)
        {
            ExLogger.LogToFile(Logger, $"{Resources.Resource.Message_UsingNuGetAPIurl}: {ExtensionConstants.NuGetApiUrl}", LogLevel.Info);
            
            var repo = PackageRepositoryFactory.Default.CreateRepository(ExtensionConstants.NuGetApiUrl);
            var packages = repo.FindPackagesById(packageId).ToList();

            return packages;
        }

        private static NuGetPackage CreateNuGetPackage(IPackage package)
        {
            return new NuGetPackage
            {
                Id = package.Id,
                Name = package.Title,
                Version = package.Version.Version,
                VersionText = package.Version.ToOriginalString(),
                XrmToolingClient = UsesXrmToolingClient(package),
                LicenseUrl = package.LicenseUrl != null
                    ? package.LicenseUrl.ToString()
                    : null
            };
        }

        private static bool UsesXrmToolingClient(IPackageMetadata package)
        {
            if (package.DependencySets?.Count() != 1)
                return false;

            foreach (var dependency in package.DependencySets.First().Dependencies)
                if (dependency.Id == ExtensionConstants.MicrosoftCrmSdkXrmToolingCoreAssembly)
                    return true;

            return false;
        }
    }
}