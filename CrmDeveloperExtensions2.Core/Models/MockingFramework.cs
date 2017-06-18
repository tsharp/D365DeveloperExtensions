using System.Collections.Generic;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class MockingFramework
    {
        public string Name { get; set; }
        public string NugetName { get; set; }
        public int CrmMajorVersion { get; set; }
    }

    public static class MockingFrameworks
    {
        public static List<MockingFramework> Frameworks => new List<MockingFramework>
        {
            new MockingFramework{CrmMajorVersion = 8, Name = "XrmUnitTest", NugetName = "XrmUnitTest.2016"},
            new MockingFramework{CrmMajorVersion = 7, Name = "XrmUnitTest", NugetName = "XrmUnitTest.2015"},
            new MockingFramework{CrmMajorVersion = 8, Name = "FakeXrmEasy", NugetName = "FakeXrmEasy.365"},
            new MockingFramework{CrmMajorVersion = 8, Name = "FakeXrmEasy", NugetName = "FakeXrmEasy.2016"},
            new MockingFramework{CrmMajorVersion = 7, Name = "FakeXrmEasy", NugetName = "FakeXrmEasy.2015"},
            new MockingFramework{CrmMajorVersion = 6, Name = "FakeXrmEasy", NugetName = "FakeXrmEasy.2013"},
            new MockingFramework{CrmMajorVersion = 5, Name = "FakeXrmEasy", NugetName = "FakeXrmEasy"}
        };
    }
}
