using System;
using System.IO;
using D365DeveloperExtensions.Core.Config;
using D365DeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EnvDTE;
using Moq;

namespace D365DeveloperExtensions.Core.Tests.Config
{
    [TestClass]
    public class ConfigFileTests
    {
        private static string _testFilepath;

        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            DirectoryInfo currentFolder = new DirectoryInfo(Environment.CurrentDirectory);
            if (currentFolder.Parent?.Parent != null)
                _testFilepath = currentFolder.Parent.Parent.FullName + "\\TestConfigFiles";
        }

        [TestMethod]
        public void SpklConfigFileExists_False()
        {
            Assert.IsFalse(ConfigFile.SpklConfigFileExists(Environment.CurrentDirectory));
        }

        [TestMethod]
        public void SpklConfigFileExists_True()
        {
            Assert.IsTrue(ConfigFile.SpklConfigFileExists(_testFilepath));
        }

        [TestMethod]
        public void GetSpklConfigFile()
        {
            //var mock = new Mock<Project>(); 
            //Project project = mock.Object;
            //project.Properties["FullName"] = "";

            //SpklConfig config = ConfigFile.GetSpklConfigFile(_testFilepath);

            //Assert.IsNotNull(config);
        }

        [TestMethod]
        public void CreateConfigFile()
        {
            //CrmDexExConfig config1 = ConfigFile.CreateConfigFile(Guid.Empty, "TestProject", _testFilepath);

            //CrmDexExConfig config2 = ConfigFile.GetConfigFile(_testFilepath);

            //Assert.AreEqual(config1.D365DevExConfigOrgMaps[0].OrganizationId, config2.D365DevExConfigOrgMaps[0].OrganizationId);
        }

        [TestMethod()]
        public void UpdateConfigFile()
        {
            //CrmDexExConfig config1 = ConfigFile.GetConfigFile(_testFilepath);

            //ConfigFile.UpdateConfigFile(_testFilepath, config1);

            //CrmDexExConfig config2 = ConfigFile.GetConfigFile(_testFilepath);

            //Assert.AreEqual(config1.D365DevExConfigOrgMaps[0].OrganizationId, config2.D365DevExConfigOrgMaps[0].OrganizationId);
        }
    }
}
