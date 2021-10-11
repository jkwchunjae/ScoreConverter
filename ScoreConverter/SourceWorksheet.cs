using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public class SourceWorksheet
    {
        private Excel.Worksheet _worksheet;
        public List<Problem> Problems;
        public List<UserData> Users;

        public SourceWorksheet(Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;

            Problems = GetProblems(worksheet);

            var userNumberList = GetUserNumberList(worksheet);
            Users = userNumberList
                .Select(x => GetUserData(x.Column, x.UserNumber, Problems, worksheet))
                .ToList();
        }

        private static List<Problem> GetProblems(Excel.Worksheet source)
        {
            var beginRow = SourceConfig.BeginRow;
            var endRow = source.GetCell(beginRow, 1).End[Excel.XlDirection.xlDown].Row;
            return Enumerable.Range(beginRow, endRow - beginRow + 1)
                .Select(row => GetSubProblem(source, row))
                .GroupBy(x => new { x.ProblemName })
                .Select(x => new Problem
                {
                    ProblemName = x.First().ProblemName,
                    SubProblems = x.ToList(),
                })
                .ToList();
        }

        private static SubProblem GetSubProblem(Excel.Worksheet source, int row)
        {
            string problemName;
            var problemCell = source.GetCell(row, SourceConfig.ProblemNameColumn);
            try
            {
                problemName = problemCell.Value2;
            }
            catch
            {
                throw new Exception($"문제 제목을 확인하세요. \n시트: {source.Name} \n셀: {problemCell.Address}");
            }

            string description;
            var descriptionCell = source.GetCell(row, SourceConfig.DescriptionColumn);
            try
            {
                description = descriptionCell.Value2;
            }
            catch
            {
                throw new Exception($"문제 설명을 확인하세요. \n시트: {source.Name} \n셀: {descriptionCell.Address}");
            }

            double score;
            var scoreCell = source.GetCell(row, SourceConfig.ScoreColumn);
            try
            {
                score = scoreCell.Value2;
                double.Parse(score.ToString());
            }
            catch
            {
                throw new Exception($"배점을 확인하세요. 숫자로 입력하세요. \n시트: {source.Name} \n셀: {scoreCell.Address}");
            }

            return new SubProblem
            {
                ProblemName = problemName,
                Description = description,
                Score = score,
                Row = row,
            };
        }

        private static List<(int Column, string UserNumber)> GetUserNumberList(Excel.Worksheet source)
        {
            var beginColumn = SourceConfig.BeginColumn;
            var endColumn = source.GetCell(SourceConfig.BeginRow - 1, beginColumn).End[XlDirection.xlToRight].Column;

            return Enumerable.Range(beginColumn, endColumn - beginColumn + 1)
                .Select(column => source.GetCell(SourceConfig.BeginRow - 1, column))
                .Select(cell =>
                {
                    var value = cell.Value2;
                    if (value is string valueStr)
                    {
                        return (cell.Column, valueStr);
                    }
                    else
                    {
                        return (cell.Column, ((object)value).ToString());
                    }
                })
                .ToList();
        }

        private static UserData GetUserData(int column, string userNumber, List<Problem> problems, Excel.Worksheet source)
        {
            var scores = problems
                .SelectMany(x => x.SubProblems)
                .Select(x => new { SubProblem = x, Cell = source.GetCell(x.Row, column) })
                .Select(x =>
                {
                    var value = x.Cell.Value2;
                    if (value is double valueDouble)
                    {
                        return new
                        {
                            x.SubProblem,
                            Score = valueDouble,
                            Cell = x.Cell,
                            Error = false,
                            ErrorMessage = string.Empty,
                        };
                    }
                    else
                    {
                        return new
                        {
                            x.SubProblem,
                            Score = 0.0,
                            Cell = x.Cell,
                            Error = true,
                            ErrorMessage = $"점수가 숫자로 변환되지 않습니다. 시트: {source.Name}, 셀: {x.Cell.Address}",
                        };
                    }
                })
                .ToList();

            var errorMessage = scores.Where(x => x.Error)
                .Select(x => x.ErrorMessage)
                .StringJoin(Environment.NewLine);

            if (scores.Any(x => x.Error))
            {
                throw new Exception(errorMessage);
            }

            return new UserData
            {
                Number = userNumber,
                Scores = scores.Select(x => (x.SubProblem, x.Score, x.Cell)).ToList(),
            };
        }
    }
}
