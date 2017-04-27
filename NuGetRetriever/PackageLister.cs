using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NuGet;
using Common.Models;

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

                results.Add(new NuGetPackage()
                {
                    Id = package.Id,
                    Name = package.Title,
                    Version = package.Version.Version,
                    VersionText = package.Version.ToOriginalString()
                });
            }

            return new List<NuGetPackage>(results.OrderByDescending(v => v.Version.ToString()));
        }
    }
}
