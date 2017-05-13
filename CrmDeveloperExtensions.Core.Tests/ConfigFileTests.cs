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
        [ClassInitialize()]
        public static void ClassInitialize(TestContext testContext)
        {
            DirectoryInfo currentFolder = new DirectoryInfo(Environment.CurrentDirectory);
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
            Config.ConfigFile.CreateConfigFile(Guid.Empty, _testFilepath);

            CrmDexExConfig config = Config.ConfigFile.GetConfigFile(_testFilepath);

            Assert.IsNotNull(config);
        }
    }
}
