using D365DeveloperExtensions.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.ObjectModel;
using System.Linq;

namespace D365DeveloperExtensions.Core.Tests.Model
{
    [TestClass]
    public class WebResourceTypeTests
    {
        [TestMethod]
        public void GetCrm5Types()
        {
            const int majorVersion = 5;
            ObservableCollection<WebResourceType> types = WebResourceTypes.GetTypes(majorVersion, false);

            var min = types.Count(t => t.CrmMinimumMajorVersion > majorVersion);
            var max = types.Count(t => t.CrmMaximumMajorVersion < majorVersion);

            Assert.IsTrue(min == 0 && max == 0 && types.Count > 0);
        }

        [TestMethod]
        public void GetCrm9Types()
        {
            const int majorVersion = 9;
            ObservableCollection<WebResourceType> types = WebResourceTypes.GetTypes(majorVersion, false);

            var min = types.Count(t => t.CrmMinimumMajorVersion > majorVersion);
            var max = types.Count(t => t.CrmMaximumMajorVersion < majorVersion);

            Assert.IsTrue(min == 0 && max == 0 && types.Count > 0);
        }

        [TestMethod]
        public void NoVersion()
        {
            const int majorVersion = 0;
            ObservableCollection<WebResourceType> types = WebResourceTypes.GetTypes(majorVersion, false);

            var min = types.Count(t => t.CrmMinimumMajorVersion > majorVersion);
            var max = types.Count(t => t.CrmMaximumMajorVersion < majorVersion);

            Assert.IsTrue(min == 0 && max == 0 && types.Count == 0);
        }
    }
}