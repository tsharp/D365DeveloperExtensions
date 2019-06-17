using System;
using System.Collections.Generic;
using D365DeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NuGetRetriever.Tests
{
    [TestClass]
    public class GetWorkflowAssembliesTests
    {
        private const string PackageId = "Microsoft.CrmSdk.Workflow";
        private static List<NuGetPackage> _packages;

        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            _packages = PackageLister.GetPackagesById(PackageId);
        }

        [TestMethod]
        public void GetWorkflowAssembliesCount()
        {
            Assert.IsTrue(_packages.Count > 0);
        }

        [TestMethod]
        public void GetWorkflowAssemblies()
        {
            Assert.AreEqual(PackageId, _packages[0].Id);
        }

    }
}
