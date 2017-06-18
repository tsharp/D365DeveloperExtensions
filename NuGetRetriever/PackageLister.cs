using CrmDeveloperExtensions2.Core.Models;
using NuGet;
using System.Collections.Generic;
using System.Linq;

namespace NuGetRetriever
{
    public static class PackageLister
    {
        public static List<NuGetPackage> GetPackagesbyId(string packageId)
        {
            IPackageRepository repo = PackageRepositoryFactory.Default.CreateRepository("https://packages.nuget.org/api/v2");
            List<IPackage> packages = repo.FindPackagesById(packageId).ToList();

            List<NuGetPackage> results = new List<NuGetPackage>();
            foreach (IPackage package in packages)
            {
                if (package.Published != null && package.Published.Value.Year == 1900)
                    continue;

                results.Add(new NuGetPackage
                {
                    Id = package.Id,
                    Name = package.Title,
                    Version = package.Version.Version,
                    VersionText = package.Version.ToOriginalString(),
                    XrmToolingClient = UsesXrmToolingClient(package)
                });
            }

            return new List<NuGetPackage>(results.OrderByDescending(v => v.Version.ToString()));
        }

        private static bool UsesXrmToolingClient(IPackageMetadata package)
        {
            if (package.DependencySets?.Count() != 1)
                return false;

            foreach (PackageDependency dependency in package.DependencySets.First().Dependencies)
                if (dependency.Id == "Microsoft.CrmSdk.XrmTooling.CoreAssembly")
                    return true;

            return false;
        }
    }
}