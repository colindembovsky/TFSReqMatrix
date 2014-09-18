using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestResultMatrix
{
    class Program
    {
        static void Main(string[] args)
        {
            var reqMatrixGenerator = new RequirementMatrixGenerator("http://localhost:8080/tfs/defaultcollection", "FabrikamFiber", "x", "Iteration 2");
            reqMatrixGenerator.Process();

            PrintMatrix(reqMatrixGenerator.Matrix);

            Console.WriteLine("Press any key to continue...");
            Console.ReadLine();
        }

        private static void PrintMatrix(RequirementMatrix matrix)
        {
            foreach (var req in matrix.Requirements)
            {
                Console.WriteLine("[{0}]: {1}", req.Id, req.Title);
                foreach (var testId in req.TestIds)
                {
                    var test = matrix.Tests.FirstOrDefault(t => t.Id == testId);
                    if (test != null)
                    {
                        Console.WriteLine("\t[{0}][{1}]: {2}", test.Id, test.MatrixState, test.Title);
                    }
                }
            }
        }
    }
}
