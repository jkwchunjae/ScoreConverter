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

        public List<Problem> Problems => Worksheet
            .Select(x => x.Problem)
            .Distinct()
            .Where(x => x != null)
            .ToList();

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
        public Excel.Worksheet Sheet;
        public Problem Problem;
        public List<string> UserNumbers;
        [JsonIgnore]
        public List<Excel.Range> UserCells;
        public List<( double Min, double Max, int Index, Excel.Range Cell)> ScoreRange;

        public TargetWorksheet(Excel.Worksheet worksheet, List<Problem> problems)
        {
            Sheet = worksheet;

            Problem = DetectProblem(worksheet, problems);
            var userData = GetUserNumberDatas(worksheet);
            UserNumbers = userData.Select(x => x.UserNumber).ToList();
            UserCells = userData.Select(x => x.Cell).ToList();
            ScoreRange = GetScoreRange(worksheet);
        }

        private static Problem DetectProblem(Excel.Worksheet worksheet, List<Problem> problems)
        {
            string problemName = worksheet.Range[TargetConfig.ProblemNameAddress].Value2;

            return problems.FirstOrDefault(x => problemName == x.ProblemName);
        }

        private static List<(string UserNumber, Excel.Range Cell)> GetUserNumberDatas(Excel.Worksheet target)
        {
            var userLeftTopCell = target.Range[TargetConfig.ScoreLeftTopAddress];
            var userLeftTopRow = userLeftTopCell.Row;

            var userNumberBeginCell = target.GetCell(userLeftTopRow, TargetConfig.UserNumberColumn);
            var userNumberEndRow = userNumberBeginCell.Offset[1, 0].Value2 is null ?
                userNumberBeginCell.Row :
                userNumberBeginCell.End[Excel.XlDirection.xlDown].Row;

            return Enumerable.Range(userLeftTopRow, userNumberEndRow - userLeftTopRow + 1)
                .Select(row => target.GetCell(row, TargetConfig.UserNumberColumn))
                .Select(cell => ((string)cell.Value2, cell))
                .ToList();
        }

        public static List<(double Min, double Max, int Index, Excel.Range Cell)> GetScoreRange(Excel.Worksheet target)
        {
            var userLeftTopCell = target.Range[TargetConfig.ScoreLeftTopAddress];
            var scoreRow = userLeftTopCell.Row - 1;
            var scoreBeginColumn = userLeftTopCell.Column;
            var scoreEndColumn = userLeftTopCell.Offset[-1, 1].Value2 is null ?
                scoreBeginColumn :
                userLeftTopCell.Offset[-1, 0].End[Excel.XlDirection.xlToRight].Column;

            return Enumerable.Range(scoreBeginColumn, scoreEndColumn - scoreBeginColumn + 1)
                .Select(column => target.GetCell(scoreRow, column))
                .Select(cell =>
                {
                    var arr = ((string)cell.Value2).Split('~');
                    var min = double.Parse(arr[0]);
                    var max = double.Parse(arr[1]);
                    var index = cell.Column - scoreBeginColumn;
                    return (min, max, index, cell);
                })
                .ToList();
        }
    }
}
