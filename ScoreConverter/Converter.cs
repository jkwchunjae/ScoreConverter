using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public static class Converter
    {
        public static bool Validate(Excel.Worksheet sourceWorksheet, Excel.Workbook targetWorkbook)
        {
            var source = new SourceWorksheet(sourceWorksheet);

            //File.WriteAllText(@"D:\source.json", JsonConvert.SerializeObject(source.Problems, Formatting.Indented), Encoding.UTF8);
            //File.WriteAllText(@"D:\score.json", JsonConvert.SerializeObject(source.Users, Formatting.Indented), Encoding.UTF8);

            var target = new TargetWorkbook(targetWorkbook, source.Problems);
            //File.WriteAllText(@"D:\target.json", JsonConvert.SerializeObject(target, Formatting.Indented), Encoding.UTF8);

            // 문제 이름이 모두 같은가?
            // TargetWorkbook 만들 때 체크하게 되어 있음.
            //var nullProblems = target.Worksheet
            //    .Where(x => x.Problem == null)
            //    .ToList();

            //if (nullProblems.Any())
            //{
            //    MessageBox.Show($"문제 이름이 다릅니다.");
            //    return false;
            //}

            // 3. 선수 수가 맞는가? 양쪽이 정확히 같아야 함.
            var targetUserList = target.Worksheet
                .GroupBy(x => x.Problem.ProblemName)
                .Select(x => x.SelectMany(e => e.UserNumbers).OrderBy(e => e).ToList())
                .First()
                .Distinct()
                .ToList();

            var sourceUserList = source.Users.Select(x => x.Number).Distinct().OrderBy(x => x).ToList();
            if (sourceUserList.Count != targetUserList.Count())
            {
                MessageBox.Show($"선수 수가 다릅니다.\n심사위원 채점표: {sourceUserList.Count} 명\n공단채점표: {targetUserList.Count} 명");
                return false;
            }

            var zipped = sourceUserList.Zip(targetUserList, (a, b) => new { Source = a, Target = b })
                .Where(x => x.Source != x.Target)
                .ToList();
            if (zipped.Any())
            {
                var sourceOnly = sourceUserList
                    .Where(x => targetUserList.Empty(e => x == e))
                    .ToList();
                var targetOnly = targetUserList
                    .Where(x => sourceUserList.Empty(e => x == e))
                    .ToList();
                MessageBox.Show($"양쪽 선수 목록이 다릅니다.\n심사위원 채점표: {sourceOnly.StringJoin("\n")}\n공단채점표: {targetOnly.StringJoin("\n")}");
                return false;
            }

            // 1. 문제수가 맞는가?
            if (source.Problems.Count != target.Problems.Count)
            {
                MessageBox.Show($"문제 수가 다릅니다.\n심사위원채점표: {source.Problems.Count} 문제\n공단채점표: {target.Problems.Count} 문제");
                return false;
            }

            // 2. 세부 항목 수가 맞는가?
            var sourceSubProblems = source.Problems
                .SelectMany(x => x.SubProblems.Select(sub => (sub.ProblemName, sub.Description)))
                .Distinct()
                .ToList();
            var targetSubProblems = target.Worksheet
                .SelectMany(x => x.ScoreRange.Select(sub => (x.Problem.ProblemName, sub.Desc)))
                .Distinct()
                .ToList();

            List<string> missingSubProblems = new List<string>();
            foreach (var targetSubProblem in targetSubProblems)
            {
                var t = targetSubProblem;
                if (sourceSubProblems.Empty(x => x.ProblemName == t.ProblemName && x.Description.StartsWith(t.Desc)))
                {
                    missingSubProblems.Add($"문제: {t.ProblemName},  항목: {t.Desc}");
                    //MessageBox.Show($"심사위원 채점표에서 \"{t.Desc}\"를 찾을 수 없습니다.");
                    //return false;
                }
            }
            if (missingSubProblems.Any())
            {
                var missing = missingSubProblems.Select(x => $"\"{x}\"")
                    .StringJoin(Environment.NewLine);
                MessageBox.Show($"심사위원 채점표에서 \n{missing}\n 를 찾을 수 없습니다.");
                return false;
            }
            foreach (var problem in targetSubProblems.Select(x => x.ProblemName))
            {
                var s = sourceSubProblems.Where(x => x.ProblemName == problem).ToList();
                var t = targetSubProblems.Where(x => x.ProblemName == problem).ToList();
                if (s.Count != t.Count)
                {
                    MessageBox.Show($"세부항목 수가 다릅니다.\n문제: {problem}\n심사위원채점표: {s.Count} 항목\n공단채점표: {t.Count} 항목");
                    return false;
                }
            }

            //foreach (var targetSheet in target.Worksheet)
            //{
            //    if (targetSheet.Problem.SubProblems.Count != targetSheet.ScoreRange.Count)
            //    {
            //        MessageBox.Show($"세부항목 수가 다릅니다.\n문제: {targetSheet.Problem.ProblemName}\n심사위원채점표: {targetSheet.Problem.SubProblems.Count} 항목\n공단채점표: {targetSheet.ScoreRange.Count} 항목");
            //        return false;
            //    }
            //}

            // 4. 점수 배점이 같은가?
            var problems = target.Worksheet
                .GroupBy(x => x.Problem.ProblemName)
                .Select(x => x.First())
                .ToList();

            foreach (var problem in problems)
            {
                var subZip = problem.Problem.SubProblems
                    .Zip(problem.ScoreRange, (a, b) => new { SubProblem = a, ScoreRange = b })
                    .ToList();

                var subDiffList = subZip.Where(x => Math.Round(x.SubProblem.Score, 3) != Math.Round(x.ScoreRange.Max, 3))
                    .ToList();

                if (subDiffList.Any())
                {
                    var subDiff = subDiffList.First();
                    MessageBox.Show($"세부항목 배점이 다릅니다.\n항목: {problem.Problem.ProblemName} - {subDiff.SubProblem.Description}\n심사위원채점표: {subDiff.SubProblem.Score}\n공단채점표: {subDiff.ScoreRange.Max}");
                    return false;
                }
            }

            // 5. 득점이 점수 구간에 있는가?

            foreach (var user in source.Users)
            {
                foreach (var score in user.Scores)
                {
                    if (score.SubProblem.Score < score.UserScore)
                    {
                        MessageBox.Show($"선수의 득점이 범위를 초과하였습니다.\n채점번호: {user.Number}\n항목: {score.SubProblem.ProblemName} - {score.SubProblem.Description}\n범위: {score.SubProblem.Score}\n득점: {score.UserScore}");
                        return false;
                    }
                }
            }

            // 실행해보자
            Execute(sourceWorksheet, targetWorkbook, execute: false);

            return true;
        }

        public static void Execute(Excel.Worksheet sourceWorksheet, Excel.Workbook targetWorkbook, bool execute)
        {
            var source = new SourceWorksheet(sourceWorksheet);
            var target = new TargetWorkbook(targetWorkbook, source.Problems);

            foreach (var targetSheet in target.Worksheet)
            {
                var targetLeftTopCell = targetSheet.Sheet.Range[TargetConfig.ScoreLeftTopAddress];
                var userDataList = targetSheet.UserNumbers.Zip(targetSheet.UserCells, (a, b) => new { UserNumber = a, Cell = b }).ToList();
                targetSheet.Sheet.Activate();

                var minRow = userDataList.Min(x => x.Cell.Row);
                var maxRow = userDataList.Max(x => x.Cell.Row);
                var rowCount = maxRow - minRow + 1;
                var subProblemCount = targetSheet.ScoreRange.Count;

                var scoreArray = new object[rowCount, subProblemCount];

                foreach (var userData in userDataList)
                {
                    var userScoreData = source.Users.First(x => x.Number == userData.UserNumber);

                    int column = 0;
                    foreach (var scoreData in targetSheet.ScoreRange)
                    {
                        var targetCell = targetSheet.Sheet.GetCell(userData.Cell.Row, targetLeftTopCell.Column + scoreData.Index);

                        var userScores = userScoreData.Scores
                            .Where(x => x.SubProblem.ProblemName == targetSheet.Problem.ProblemName)
                            .Where(x => x.SubProblem.Description.StartsWith(scoreData.Desc))
                            .ToList();

                        if (userScores.Empty())
                        {
                            throw new Exception($"{targetCell.Address} 심사위원 채점표에서 \"{targetCell.Offset[-2, 0].Value2}\"를 찾을 수 없습니다.");
                        }

                        var userScore = userScores.First().UserScore;

                        // targetCell.Value2 = userScore.ToString();
                        scoreArray[userData.Cell.Row - minRow, column] = userScore.ToString();
                        column++;
                    }
                }

                if (execute)
                {
                    var minCell = targetLeftTopCell;
                    var maxCell = minCell.Offset[rowCount - 1, subProblemCount - 1];
                    targetSheet.Sheet.Range[minCell, maxCell].Value2 = scoreArray;
                }
            }
        }
    }
}
