using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestResultMatrix
{
    public class Requirement : WorkItemInfo
    {
        public List<int> TestIds { get; set; }

        public Requirement()
        {
            TestIds = new List<int>();
        }
    }
}
