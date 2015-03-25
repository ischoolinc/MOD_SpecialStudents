using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    public class StudentAbsence
    {
        /// <summary>
        /// 全勤學生
        /// </summary>
        public List<string> StudentNoAbsenceList { get; set; }

        /// <summary>
        /// 有缺曠記錄的學生
        /// </summary>
        public List<string> StudentAbsenceList { get; set; }

        public StudentAbsence()
        {
            StudentNoAbsenceList = new List<string>();
            StudentAbsenceList = new List<string>();
        }
    }
}
