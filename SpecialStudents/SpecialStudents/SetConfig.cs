using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    public class SetConfig
    {
        /// <summary>
        /// 模式
        /// </summary>
        public SelectMode _selectMode { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public int _SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int _Semester { get; set; }

        /// <summary>
        /// 開始日期
        /// </summary>
        public DateTime StartDate { get; set; }

        /// <summary>
        /// 結束日期
        /// </summary>
        public DateTime EndDate { get; set; }

        /// <summary>
        /// 學生
        /// </summary>
        public List<StudentRecord> _StudentList { get; set; }

        /// <summary>
        /// 取得學生清單ID
        /// </summary>
        /// <returns></returns>
        public List<string> GetStudentIdList()
        {
            List<string> list = new List<string>();
            if (_StudentList != null)
            {
                foreach (StudentRecord each in _StudentList)
                {
                    if (!list.Contains(each.ID))
                    {
                        list.Add(each.ID);
                    }
                }
            }
            return list;

        }
    }
}
