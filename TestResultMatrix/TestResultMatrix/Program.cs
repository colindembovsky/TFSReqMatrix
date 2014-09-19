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
			try
			{
				if (args.Count() < 2 || args.Count() > 3)
				{
					Console.WriteLine("Usage: TestResultMatrix.exe tpcUrl teamProjectName [requirementQueryName]");
					Console.WriteLine();
					Console.WriteLine("  tpcUrl: url to Team Project collection - e.g. http://localhost:8080/tfs/defaultcollection");
					Console.WriteLine("  teamProjectName: name of Team Project - e.g. FabFiber");
					Console.WriteLine("  requirementQueryName: (optional) flat-list query of requirements.");
					Console.WriteLine();
					Console.WriteLine("If you do not specify requirementQueryName, the tool will get all work items in the requirement category");
					Console.WriteLine();
					return;
				}

				string reqQuery = null;
				if (args.Count() == 3)
				{
					reqQuery = args[2];
				}

				var reqMatrixGenerator = new RequirementMatrixGenerator(args[0], args[1], reqQuery);
				reqMatrixGenerator.Process();

				//PrintMatrix(reqMatrixGenerator.Matrix);

				var excel = new MatrixExcel(reqMatrixGenerator.Matrix);
				excel.GenerateMatrixSheet();
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.ToString());
			}
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

            Console.WriteLine("Press any key to continue...");
            Console.ReadLine();
        }
    }
}
