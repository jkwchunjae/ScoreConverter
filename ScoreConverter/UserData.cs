using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreConverter
{
    public class UserData
    {
        public string Number;
        public List<(SubProblem SubProblem, double UserScore)> Scores;
    }
}
