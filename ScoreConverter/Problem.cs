using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreConverter
{
    public class SubProblem
    {
        public string ProblemName;
        public int SubNo;
        public double Score;
        public string Description;
        public int Row;
    }

    public class Problem
    {
        public string ProblemName;
        public List<SubProblem> SubProblems;

        public SubProblem GetSubProblem(int row)
        {
            return SubProblems.FirstOrDefault(x => x.Row == row);
        }
    }
}
