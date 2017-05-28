using System;
using System.IO;
using CrmDeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CrmDeveloperExtensions.Core.Tests
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
            Assert.IsFalse(Config.ConfigFile.ConfigFileExists(Environment.CurrentDirectory));
        }

        [TestMethod]
        public void ConfigFileExists_True()
        {
            Assert.IsTrue(Config.ConfigFile.ConfigFileExists(_testFilepath));
        }

        [TestMethod]
        public void GetConfigFile()
        {
            CrmDexExConfig config = Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.IsNotNull(config);
        }

        [TestMethod]
        public void CreateConfigFile()
        {
            CrmDexExConfig config1 = Config.ConfigFile.CreateConfigFile(Guid.Empty, "TestProject", _testFilepath);

            CrmDexExConfig config2 = Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.AreEqual(config1.CrmDevExConfigOrgMaps[0].OrganizationId, config2.CrmDevExConfigOrgMaps[0].OrganizationId);
        }
    }
}
