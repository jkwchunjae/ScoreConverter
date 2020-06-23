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

            Users = GetUserNumberList(worksheet)
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
            try { problemName = source.GetCell(row, 1).Value2; }
            catch { MessageBox.Show($"Error on MakeSubProblem. \nSheet: {source.Name} \nRow: {row}"); throw; }
            string description;
            try { description = source.GetCell(row, 5).Value2; }
            catch { MessageBox.Show($"Error on MakeSubProblem. \nSheet: {source.Name} \nRow: {row}"); throw; }
            double subNo;
            try { subNo = source.GetCell(row, 2).Value2; }
            catch { MessageBox.Show($"Error on MakeSubProblem. \nSheet: {source.Name} \nRow: {row}"); throw; }
            double score;
            try { score = source.GetCell(row, 4).Value2; }
            catch { MessageBox.Show($"Error on MakeSubProblem. \nSheet: {source.Name} \nRow: {row}"); throw; }

            return new SubProblem
            {
                ProblemName = problemName,
                Description = description,
                SubNo = (int)subNo,
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
                    if (value is string)
                    {
                        return (cell.Column, (string)value);
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
                    var score = 0.0;
                    if (value is double)
                    {
                        score = (double)value;
                    }
                    return (x.SubProblem, score);
                })
                .ToList();

            return new UserData
            {
                Number = userNumber,
                Scores = scores,
            };
        }
    }
}
