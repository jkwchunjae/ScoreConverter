using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreConverter
{
    public static class SourceConfig
    {
        /// <summary>
        /// 첫 문제 번호가 시작하는 행
        /// 다른 말로.. 제목 다음 행
        /// </summary>
        public static int BeginRow { get; } = 5;


        /// <summary>
        /// (값: 1) 문제이름(항목명)이 적혀있는 열
        /// TargetConfig.ProblemNameAddress 에 적혀있는 값이 있어야 한다.
        /// 쉽게 말해서 공단에서 다운받은 엑셀 파일의 ProblemNameAddress(F4) 주소에 써있는 항목명이 써있어야 한다.
        /// </summary>
        public static int ProblemNameColumn { get; } = 1;


        /// <summary>
        /// (값: 2) 설명이 써있는 열
        /// </summary>
        public static int DescriptionColumn { get; } = 2;


        /// <summary>
        /// (값: 3) 세부 항목의 배점이 적혀있는 열
        /// </summary>
        public static int ScoreColumn { get; } = 3;


        /// <summary>
        /// (값: 4) 1번 선수가 있는 열
        /// </summary>
        public static int BeginColumn { get; } = 4;
    }

    public static class TargetConfig
    {
        /// <summary> 문제이름 (항목명)의 셀 주소 </summary>
        public static string ProblemNameAddress { get; } = "F4";

        /// <summary> 점수를 입력하는 셀 주소 </summary>
        public static string ScoreLeftTopAddress { get; } = "D10";

        /// <summary> 채점비번호 열 </summary>
        public static int UserNumberColumn { get; } = 2;
    }
}
