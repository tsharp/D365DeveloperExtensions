using System;
using System.IO;
using D365DeveloperExtensions.Core.Config;
using D365DeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace D365DeveloperExtensions.Core.Tests.Config
{
    [TestClass]
    public class ConfigFileTests
    {
        //private static string _testFilepath;
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            //DirectoryInfo currentFolder = new DirectoryInfo(Environment.CurrentDirectory);
            //if (currentFolder.Parent?.Parent != null)
            //    _testFilepath = currentFolder.Parent.Parent.FullName + "\\TestConfigFiles";
        }

        [TestMethod]
        public void ConfigFileExists_False()
        {
            //Assert.IsFalse(ConfigFile.ConfigFileExists(Environment.CurrentDirectory));
        }

        [TestMethod]
        public void ConfigFileExists_True()
        {
            //Assert.IsTrue(ConfigFile.ConfigFileExists(_testFilepath));
        }

        [TestMethod]
        public void GetConfigFile()
        {
            //CrmDexExConfig config = ConfigFile.GetConfigFile(_testFilepath);

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
