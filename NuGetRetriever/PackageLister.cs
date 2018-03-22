using D365DeveloperExtensions.Core;
using D365DeveloperExtensions.Core.Models;
using NuGet;
using System.Collections.Generic;
using System.Linq;

namespace NuGetRetriever
{

    public static class PackageLister
    {
        public static List<NuGetPackage> GetPackagesbyId(string packageId)
        {
            var packages = GetPackages(packageId);

            List<NuGetPackage> results = new List<NuGetPackage>();
            foreach (IPackage package in packages)
            {
                if (package.Published != null && package.Published.Value.Year == 1900)
                    continue;

                results.Add(CreateNuGetPackage(package));
            }

            return new List<NuGetPackage>(results.OrderByDescending(v => v.Version.ToString()));
        }

        private static List<IPackage> GetPackages(string packageId)
        {
            IPackageRepository repo = PackageRepositoryFactory.Default.CreateRepository("https://packages.nuget.org/api/v2");
            List<IPackage> packages = repo.FindPackagesById(packageId).ToList();

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

            foreach (PackageDependency dependency in package.DependencySets.First().Dependencies)
                if (dependency.Id == ExtensionConstants.MicrosoftCrmSdkXrmToolingCoreAssembly)
                    return true;

            return false;
        }
    }
}