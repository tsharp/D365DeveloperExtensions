using System;
using System.IO;
using CrmDeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CrmDeveloperExtensions.Core.Tests.Config
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
        public void ConfigFileExists_False()
        {
            Assert.IsFalse(Core.Config.ConfigFile.ConfigFileExists(Environment.CurrentDirectory));
        }

        [TestMethod]
        public void ConfigFileExists_True()
        {
            Assert.IsTrue(Core.Config.ConfigFile.ConfigFileExists(_testFilepath));
        }

        [TestMethod]
        public void GetConfigFile()
        {
            CrmDexExConfig config = Core.Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.IsNotNull(config);
        }

        [TestMethod]
        public void CreateConfigFile()
        {
            CrmDexExConfig config1 = Core.Config.ConfigFile.CreateConfigFile(Guid.Empty, "TestProject", _testFilepath);

            CrmDexExConfig config2 = Core.Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.AreEqual(config1.CrmDevExConfigOrgMaps[0].OrganizationId, config2.CrmDevExConfigOrgMaps[0].OrganizationId);
        }

        [TestMethod()]
        public void UpdateConfigFile()
        {
            CrmDexExConfig config1 = Core.Config.ConfigFile.GetConfigFile(_testFilepath);

            Core.Config.ConfigFile.UpdateConfigFile(_testFilepath, config1);

            CrmDexExConfig config2 = Core.Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.AreEqual(config1.CrmDevExConfigOrgMaps[0].OrganizationId, config2.CrmDevExConfigOrgMaps[0].OrganizationId);
        }
    }
}
