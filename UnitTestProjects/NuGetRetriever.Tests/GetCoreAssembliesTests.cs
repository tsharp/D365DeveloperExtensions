using System;
using System.Collections.Generic;
using D365DeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NuGetRetriever.Tests
{
    [TestClass]
    public class GetCoreAssembliesTests
    {
        private const string PackageId = "Microsoft.CrmSdk.CoreAssemblies";
        private static List<NuGetPackage> _packages;

        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            _packages = PackageLister.GetPackagesById(PackageId);
        }

        [TestMethod]
        public void GetCoreAssembliesCount()
        {
            Assert.IsTrue(_packages.Count > 0);
        }

        [TestMethod]
        public void GetCoreAssemblies()
        {
            Assert.AreEqual(PackageId, _packages[0].Id);
        }

    }
}
