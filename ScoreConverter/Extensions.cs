using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public static class Extensions
    {
        public static IEnumerable<Excel.Workbook> GetWorkbooks(this Excel.Application application)
        {
            for (var i = 1; i <= application.Workbooks.Count; i++)
            {
                yield return application.Workbooks[i];
            }
        }

        public static IEnumerable<Excel.Workbook> GetWorkbooks(this Excel.Application application, Func<Excel.Workbook, bool> func)
        {
            return application.GetWorkbooks()
                .Where(func);
        }

        public static bool TryGetWorkbook(this Excel.Application application, Func<Excel.Workbook, bool> func, out Excel.Workbook workbook)
        {
            var find = application.GetWorkbooks(func).ToList();
            if (find.Any())
            {
                workbook = find.First();
                return true;
            }
            else
            {
                workbook = null;
                return false;
            }
        }

        public static IEnumerable<Excel.Worksheet> GetWorksheets(this Excel.Workbook workbook)
        {
            for (var i = 1; i <= workbook.Sheets.Count; i++)
            {
                yield return workbook.Sheets[i];
            }
        }

        public static IEnumerable<Excel.Worksheet> GetWorksheets(this Excel.Workbook workbook, Func<Excel.Worksheet, bool> func)
        {
            return workbook.GetWorksheets()
                .Where(func);
        }

        public static bool TryGetWorksheet(this Excel.Workbook workbook, Func<Excel.Worksheet, bool> func, out Excel.Worksheet worksheet)
        {
            var find = workbook.GetWorksheets(func).ToList();
            if (find.Any())
            {
                worksheet = find.First();
                return true;
            }
            else
            {
                worksheet = null;
                return false;
            }
        }

        public static Excel.Range GetCell(this Excel.Worksheet worksheet, int row, int column)
        {
            return worksheet.Cells[row, column];
        }

        public static string StringJoin(this IEnumerable<string> source, string separator)
        {
            return string.Join(separator, source);
        }

        public static bool Empty<T>(this IEnumerable<T> source)
        {
            return !source.Any();
        }

        public static bool Empty<T>(this IEnumerable<T> source, Func<T, bool> predicate)
        {
            return !source.Any(predicate);
        }
    }
}
