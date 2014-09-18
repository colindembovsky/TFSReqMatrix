using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestResultMatrix
{
    public class RequirementMatrix
    {
        public List<WorkItemInfo> Tests { get; set; }

        public List<Requirement> Requirements { get; set; }

        public RequirementMatrix()
        {
            Tests = new List<WorkItemInfo>();
            Requirements = new List<Requirement>();
        }
    }
}
