using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            return true;
        }
    }
}
