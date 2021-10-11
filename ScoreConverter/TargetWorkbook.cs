using Newtonsoft.Json;
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

        public List<string> Problems => Worksheet
            .Select(x => x.ProblemName)
            .Distinct()
            .Where(x => x != null)
            .ToList();

        public TargetWorkbook(Excel.Workbook workbook)
        {
            var sheets = workbook.GetWorksheets();

            Worksheet = sheets
                .Select(x => new TargetWorksheet(x))
                .ToList();
        }
    }

    public class TargetWorksheet
    {
        public Excel.Worksheet Sheet { get; private set; }
        public string ProblemName { get; private set; }
        public List<(string Number, Excel.Range Cell)> UserData { get; private set; }
        public List<( double Min, double Max, int Column, string Desc, Excel.Range Cell)> SubProblems { get; private set; }

        public TargetWorksheet(Excel.Worksheet worksheet)
        {
            Sheet = worksheet;

            ProblemName = GetProblemName(worksheet);
            UserData = GetUserNumberDatas(worksheet);
            SubProblems = GetScoreRange(worksheet);
        }

        private string GetProblemName(Excel.Worksheet worksheet)
        {
            string problemName = worksheet.Range[TargetConfig.ProblemNameAddress].Value2;

            return problemName;
        }

        private static List<(string Number, Excel.Range Cell)> GetUserNumberDatas(Excel.Worksheet target)
        {
            var userLeftTopCell = target.Range[TargetConfig.ScoreLeftTopAddress];
            var userLeftTopRow = userLeftTopCell.Row;

            var userNumberBeginCell = target.GetCell(userLeftTopRow, TargetConfig.UserNumberColumn);
            var userNumberEndRow = userNumberBeginCell.Offset[1, 0].Value2 is null ? // 시트에 선수가 한명만 있는 경우
                userNumberBeginCell.Row :
                userNumberBeginCell.End[Excel.XlDirection.xlDown].Row;

            return Enumerable.Range(userLeftTopRow, userNumberEndRow - userLeftTopRow + 1)
                .Select(row => target.GetCell(row, TargetConfig.UserNumberColumn))
                .Select(cell => ((string)cell.Value2, cell))
                .ToList();
        }

        public static List<(double Min, double Max, int Column, string Description, Excel.Range Cell)> GetScoreRange(Excel.Worksheet target)
        {
            var userLeftTopCell = target.Range[TargetConfig.ScoreLeftTopAddress];
            var scoreRow = userLeftTopCell.Row - 1;
            var scoreBeginColumn = userLeftTopCell.Column;
            var scoreEndColumn = userLeftTopCell.Offset[-1, 1].Value2 is null ? // 시트에 세부 문제가 1개인 경우
                scoreBeginColumn :
                userLeftTopCell.Offset[-1, 0].End[Excel.XlDirection.xlToRight].Column;

            return Enumerable.Range(scoreBeginColumn, scoreEndColumn - scoreBeginColumn + 1)
                .Select(column => target.GetCell(scoreRow, column))
                .Where(cell => cell.Offset[-1, 0].Value2 != "총계")
                .Select(cell =>
                {
                    try
                    {
                        var arr = ((string)cell.Value2).Split('~');
                        var min = double.Parse(arr[0]);
                        var max = double.Parse(arr[1]);
                        string desc = (string)cell.Offset[-1, 0].Value2;
                        return (min, max, cell.Column, desc, cell);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"{target.Name}, {cell.Address}");
                    }
                })
                .ToList();
        }
    }
}
