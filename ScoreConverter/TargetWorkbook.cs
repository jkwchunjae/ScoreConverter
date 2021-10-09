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
        [JsonIgnore]
        public Excel.Worksheet Sheet;
        public string ProblemName => Problem?.ProblemName ?? _problemName;
        public Problem Problem;
        public List<string> UserNumbers;
        [JsonIgnore]
        public List<Excel.Range> UserCells;
        public List<( double Min, double Max, int Index, string Desc, Excel.Range Cell)> ScoreRange;

        private string _problemName = string.Empty;

        public TargetWorksheet(Excel.Worksheet worksheet, List<Problem> problems)
        {
            Sheet = worksheet;

            Problem = DetectProblem(worksheet, problems);
            var userData = GetUserNumberDatas(worksheet);
            UserNumbers = userData.Select(x => x.UserNumber).ToList();
            UserCells = userData.Select(x => x.Cell).ToList();
            ScoreRange = GetScoreRange(worksheet);
        }

        private Problem DetectProblem(Excel.Worksheet worksheet, List<Problem> problems)
        {
            string problemName = worksheet.Range[TargetConfig.ProblemNameAddress].Value2;

            if (problems == null)
            {
                _problemName = problemName;
                return null;
            }

            var problem = problems.FirstOrDefault(x => x.ProblemName == problemName);
            if (problem == null)
            {
                throw new Exception($"문제를 찾을 수 없습니다.\n시트: {worksheet.Name}\n문제 이름: {problemName}");
            }
            else
            {
                return problem;
            }
        }

        private static List<(string UserNumber, Excel.Range Cell)> GetUserNumberDatas(Excel.Worksheet target)
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

        public static List<(double Min, double Max, int Index, string Description, Excel.Range Cell)> GetScoreRange(Excel.Worksheet target)
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
                        var index = cell.Column - scoreBeginColumn;
                        string desc = (string)cell.Offset[-1, 0].Value2;
                        return (min, max, index, desc, cell);
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
