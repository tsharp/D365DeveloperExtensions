using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;

namespace CrmDeveloperExtensions.Core.Connection
{
    public class SolutionProjectRemovedEventArgs
    {
        public Project Project { get; set; }
    }
}
