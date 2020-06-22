using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public class TargetWorkbook
    {
        public List<TargetWorksheet> Worksheet;

        public TargetWorkbook(Excel.Workbook workbook, List<Problem> problems)
        {
            var sheets = workbook.GetWorksheets();

            Worksheet = sheets
                .Select(x => new TargetWorksheet(x, problems))
                .ToList();
        }
    }

    public class TargetWorksheet
    {
        Excel.Worksheet _sheet;
        public Problem Problem;

        public TargetWorksheet(Excel.Worksheet worksheet, List<Problem> problems)
        {
            _sheet = worksheet;

            Problem = DetectProblem(worksheet, problems);
        }

        private static Problem DetectProblem(Excel.Worksheet worksheet, List<Problem> problems)
        {
            string problemName = worksheet.Range[TargetConfig.ProblemNameAddress].Value2;

            return problems.FirstOrDefault(x => problemName.Contains(x.ProblemName));
        }
    }
}
