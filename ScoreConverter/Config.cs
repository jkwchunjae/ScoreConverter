using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreConverter
{
    public static class SourceConfig
    {
        public static readonly int BeginRow = 5; // 문제1 이 시작하는 행
        public static readonly int BeginColumn = 6; // 1번 선수 열
        public static readonly int ScoreColumn = 4; // 배점 열
    }

    public static class TargetConfig
    {
        /// <summary> 문제이름 (항목명)의 셀 주소 </summary>
        public static readonly string ProblemNameAddress = "F4";

        /// <summary> 점수를 입력하는 셀 주소 </summary>
        public static readonly string ScoreLeftTopAddress = "D10";

        /// <summary> 채점비번호 열 </summary>
        public static readonly int UserNumberColumn = 2;
    }
}
