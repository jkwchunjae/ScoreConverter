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
    public static class Validator
    {
        public static bool Validate(Excel.Worksheet sourceWorksheet, Excel.Workbook targetWorkbook)
        {
            var source = new SourceWorksheet(sourceWorksheet);

            File.WriteAllText(@"D:\source.json", JsonConvert.SerializeObject(source.Problems, Formatting.Indented), Encoding.UTF8);
            File.WriteAllText(@"D:\score.json", JsonConvert.SerializeObject(source.Users, Formatting.Indented), Encoding.UTF8);

            var target = new TargetWorkbook(targetWorkbook, source.Problems);
            File.WriteAllText(@"D:\target.json", JsonConvert.SerializeObject(target, Formatting.Indented), Encoding.UTF8);

            // 문제 이름이 모두 같은가?
            var nullProblems = target.Worksheet
                .Where(x => x.Problem == null)
                .ToList();

            if (nullProblems.Any())
            {
                MessageBox.Show($"문제 이름이 다릅니다.");
                return false;
            }

            // 3. 선수 수가 맞는가? 양쪽이 정확히 같아야 함.
            var targetUserList = target.Worksheet
                .GroupBy(x => x.Problem.ProblemName)
                .Select(x => x.SelectMany(e => e.UserNumbers).OrderBy(e => e).ToList())
                .First();

            var sourceUserList = source.Users.Select(x => x.Number).OrderBy(x => x).ToList();
            if (sourceUserList.Count != targetUserList.Count)
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
            foreach (var targetSheet in target.Worksheet)
            {
                if (targetSheet.Problem.SubProblems.Count != targetSheet.ScoreRange.Count)
                {
                    MessageBox.Show($"세부항목 수가 다릅니다.\n문제: {targetSheet.Problem.ProblemName}\n심사위원채점표: {targetSheet.Problem.SubProblems.Count} 항목\n공단채점표: {targetSheet.ScoreRange.Count} 항목");
                    return false;
                }
            }

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

                var subDiffList = subZip.Where(x => x.SubProblem.Score != x.ScoreRange.Max)
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

            return true;
        }
    }
}
