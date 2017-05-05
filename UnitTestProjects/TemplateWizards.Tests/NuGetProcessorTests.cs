using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateWizards;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TemplateWizards.Enums;

namespace TemplateWizards.Tests
{
    [TestClass()]
    public class NuGetProcessorTests
    {
        [TestMethod()]
        public void DetermineClientTypeTest_ValidXrmTooling1()
        {
            string expected = "Microsoft.CrmSdk.XrmTooling.CoreAssembly";
            string result = NuGetProcessor.DetermineClientType("8.2.0");
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void DetermineClientTypeTest_ValidXrmTooling2()
        {
            string expected = "Microsoft.CrmSdk.XrmTooling.CoreAssembly";
            string result = NuGetProcessor.DetermineClientType("6.1.0");
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void DetermineClientTypeTest_ValidXrmTooling3Preview()
        {
            string expected = "Microsoft.CrmSdk.XrmTooling.CoreAssembly";
            string result = NuGetProcessor.DetermineClientType("6.1.0-preview");
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void DetermineClientTypeTest_ValidExtensions1()
        {
            string expected = "Microsoft.CrmSdk.Extensions";
            string result = NuGetProcessor.DetermineClientType("5.0.17");
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void DetermineClientTypeTest_ValidExtensions2()
        {
            string expected = "Microsoft.CrmSdk.Extensions";
            string result = NuGetProcessor.DetermineClientType("6.0.4");
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void GetNextPackageTestNextWorkflow1()
        {
            int expected = 2;
            int result = NuGetProcessor.GetNextPackage(PackageValue.Core, true, true);
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void GetNextPackageTestNextClient()
        {
            int expected = 3;
            int result = NuGetProcessor.GetNextPackage(PackageValue.Workflow, false, true);
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void GetNextPackageTestClose1()
        {
            int expected = 0;
            int result = NuGetProcessor.GetNextPackage(PackageValue.Core, false, false);
            Assert.AreEqual(expected, result);
        }

        [TestMethod()]
        public void GetNextPackageTestClose2()
        {
            int expected = 0;
            int result = NuGetProcessor.GetNextPackage(PackageValue.Client, true, true);
            Assert.AreEqual(expected, result);
        }
    }
}