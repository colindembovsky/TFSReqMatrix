using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestResultMatrix
{
    public class MatrixExcel
    {
        private const int TestCaseColOffset = 8;

        public Excel.Application ExcelApp { get; private set; }

        public Excel.Workbook Workbook { get; private set; }

        public Excel.Worksheet Worksheet { get; private set; }

        public RequirementMatrix Matrix { get; private set; }

        public MatrixExcel(RequirementMatrix matrix)
        {
            ExcelApp = new Excel.Application();
            ExcelApp.Visible = true;
            Workbook = ExcelApp.Workbooks.Add(1);
            Worksheet = (Excel.Worksheet)Workbook.Sheets[1];
            Matrix = matrix;
        }

        public void GenerateMatrixSheet()
        {
            AddIntersectionStyles();
            GenerateHeaders();
            GenerateIntersections();

            StyleSheet();
        }

        private void AddIntersectionStyles()
        {
            var pass = Workbook.Styles.Add("Passed");
            pass.Interior.Color = XlRgbColor.rgbGreen;

            var fail = Workbook.Styles.Add("Failed");
            fail.Interior.Color = XlRgbColor.rgbRed;

            var blocked = Workbook.Styles.Add("Blocked");
            blocked.Interior.Color = XlRgbColor.rgbOrange;

            var notRun = Workbook.Styles.Add("Not run");
            notRun.Interior.Color = XlRgbColor.rgbLightBlue;
        }

        private void GenerateTotals(Requirement req, int row)
        {
            var totals = from t in Matrix.Tests
                         where req.TestIds.Contains(t.Id)
                         group t by t.MatrixState into res
                         select new KeyValuePair<string, int>(res.Key, res.Count());

            var stateCol = 4;
            new[] { "Passed", "Failed", "Blocked", "Not run" }.ToList().ForEach(state =>
                {
                    var total = GetTotal(state, totals);
                    if (total == 0)
                    {
                        Worksheet.Cells[row, stateCol++] = "";
                    }
                    else
                    {
                        Worksheet.Cells[row, stateCol++] = total;
                    }
                });
        }

        private int GetTotal(string state, IEnumerable<KeyValuePair<string, int>> totals)
        {
            if (totals.Any(p => p.Key == state))
            {
                var pair = totals.Single(p => p.Key == state);
                return pair.Value;
            }
            return 0;
        }

        private void GenerateHeaders()
        {
            int row = 4;
            foreach (var req in Matrix.Requirements.OrderBy(r => r.Id))
            {
                Worksheet.Cells[row, 1] = req.Id;
                Worksheet.Cells[row, 2] = req.Title;
                Worksheet.Cells[row, 3] = req.MatrixState;
                GenerateTotals(req, row++);
            }

            int col = TestCaseColOffset;
            foreach (var test in Matrix.Tests.OrderBy(t => t.Id))
            {
                Worksheet.Cells[2, col++] = string.Format("[{0}] {1}", test.Id, test.Title);
            }
        }

        private void GenerateIntersections()
        {
            var row = 4;
            foreach (var req in Matrix.Requirements.OrderBy(r => r.Id))
            {
                var col = TestCaseColOffset;
                foreach(var test in Matrix.Tests)
                {
                    if (req.TestIds.Contains(test.Id))
                    {
                        Worksheet.Cells[row, col].Style = test.MatrixState;
                    }
                    col++;
                }
                row++;
            }
        }

        private void StyleSheet()
        {
            // add headers
            Worksheet.Cells[3, 1] = "ID";
            Worksheet.Cells[3, 2] = "Title";
            Worksheet.Cells[3, 3] = "State";
            Worksheet.Cells[3, 4] = "Passed";
            Worksheet.Cells[3, 5] = "Failed";
            Worksheet.Cells[3, 6] = "Blocked";
            Worksheet.Cells[3, 7] = "Not Run";

            // autofit the columns
            for (int i = 1; i <= 3; i++)
            {
                Worksheet.Cells[1, i].EntireColumn.AutoFit();
            }

            for (int i = 4; i < TestCaseColOffset; i++)
            {
                Worksheet.Cells[1, i].EntireColumn.ColumnWidth = 7d;
            }

            for (int i = TestCaseColOffset; i < TestCaseColOffset + Matrix.Tests.Count; i++)
            {
                Worksheet.Cells[1, i].EntireColumn.ColumnWidth = 5d;
            }

            // format test column to orient to 45 degrees with borders
            var testHeaderRange = Worksheet.Range[Worksheet.Cells[2, TestCaseColOffset], Worksheet.Cells[2, TestCaseColOffset + Matrix.Tests.Count]];
            testHeaderRange.Orientation = 45;
            testHeaderRange.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            testHeaderRange.Borders[XlBordersIndex.xlInsideVertical].Weight = 1d;
            testHeaderRange.Borders[XlBordersIndex.xlInsideVertical].Color = XlRgbColor.rgbBlack;
            testHeaderRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            testHeaderRange.Borders[XlBordersIndex.xlEdgeLeft].Weight = 1d;
            testHeaderRange.Borders[XlBordersIndex.xlEdgeLeft].Color = XlRgbColor.rgbBlack;

            // format headers
            var headerRange = Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, TestCaseColOffset - 1]];
            headerRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, null, XlRgbColor.rgbBlack);
            headerRange.Interior.Color = XlRgbColor.rgbLightGray;

            // borders for reqs, summary and matrix
            var reqRange = Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3 + Matrix.Requirements.Count, 3]];
            reqRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, null, XlRgbColor.rgbBlack);

            var summaryRange = Worksheet.Range[Worksheet.Cells[3, 4], Worksheet.Cells[3 + Matrix.Requirements.Count, 7]];
            summaryRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, null, XlRgbColor.rgbBlack);

            var matrixRange = Worksheet.Range[Worksheet.Cells[3, TestCaseColOffset], Worksheet.Cells[3 + Matrix.Requirements.Count, TestCaseColOffset + Matrix.Tests.Count - 1]];
            matrixRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, null, XlRgbColor.rgbBlack);

            // color format for summary section
            var passRange = Worksheet.Range[Worksheet.Cells[4, 4], Worksheet.Cells[4 + Matrix.Requirements.Count, 4]];
            var passBar = passRange.FormatConditions.AddDatabar();
            passBar.BarColor.Color = XlRgbColor.rgbGreen;

            var failRange = Worksheet.Range[Worksheet.Cells[4, 5], Worksheet.Cells[4 + Matrix.Requirements.Count, 5]];
            var failBar = failRange.FormatConditions.AddDatabar();
            failBar.BarColor.Color = XlRgbColor.rgbRed;

            var blockedRange = Worksheet.Range[Worksheet.Cells[4, 6], Worksheet.Cells[4 + Matrix.Requirements.Count, 6]];
            var blockedBar = blockedRange.FormatConditions.AddDatabar();
            blockedBar.BarColor.Color = XlRgbColor.rgbOrange;

            var notRunRange = Worksheet.Range[Worksheet.Cells[4, 7], Worksheet.Cells[4 + Matrix.Requirements.Count, 7]];
            var notRunBar = notRunRange.FormatConditions.AddDatabar();
            notRunBar.BarColor.Color = XlRgbColor.rgbLightBlue;

            // add filters to header row
            var allHeaderRange = Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, TestCaseColOffset + Matrix.Tests.Count - 1]];
            allHeaderRange.AutoFilter();
        }
    }
}
