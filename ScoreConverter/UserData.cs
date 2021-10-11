using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public class UserData
    {
        public string Number;
        public List<(SubProblem SubProblem, double UserScore, Excel.Range Cell)> Scores;
    }
}
