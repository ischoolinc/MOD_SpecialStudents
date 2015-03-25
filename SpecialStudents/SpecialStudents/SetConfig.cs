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

        #region 獎勵特殊表現

        /// <summary>
        /// 大功轉小功
        /// </summary>
        public int Meritab { get; set; }

        /// <summary>
        /// 小功轉嘉獎
        /// </summary>
        public int Meritbc { get; set; }

        /// <summary>
        /// 大過轉小過
        /// </summary>
        public int Demeritab { get; set; }

        /// <summary>
        /// 小過轉警告
        /// </summary>
        public int Demeritbc { get; set; }

        /// <summary>
        /// 使用者設定之統計基數
        /// </summary>
        public int Meritwant { get; set; }

        /// <summary>
        /// 使用者設定之統計基數
        /// </summary>
        public int Demeritwant { get; set; }

        /// <summary>
        /// 取得獎勵換算原則
        /// </summary>
        public void GetMerReduce()
        {
            MeritDemeritReduceRecord record = K12.Data.MeritDemeritReduce.Select();
            Meritab = record.MeritAToMeritB.HasValue ? record.MeritAToMeritB.Value : 1;
            Meritbc = record.MeritBToMeritC.HasValue ? record.MeritBToMeritC.Value : 1;

            Demeritab = record.DemeritAToDemeritB.HasValue ? record.DemeritAToDemeritB.Value : 1;
            Demeritbc = record.DemeritBToDemeritC.HasValue ? record.DemeritBToDemeritC.Value : 1;

            //計算出基數
            Meritwant = (MeritCountA * Meritab * Meritbc) + (MeritCountB * Meritbc) + MeritCountC;

            Demeritwant = (DemeritCountA * Meritab * Meritbc) + (DemeritCountB * Meritbc) + DemeritCountC;
        }

        #endregion

        #region 使用者設定

        /// <summary>
        /// 使用者設定 - 大功
        /// </summary>
        public int MeritCountA { get; set; }

        /// <summary>
        /// 使用者設定 - 小功
        /// </summary>
        public int MeritCountB { get; set; }

        /// <summary>
        /// 使用者設定 - 嘉獎
        /// </summary>
        public int MeritCountC { get; set; }

        /// <summary>
        /// 使用者設定 - 大過
        /// </summary>
        public int DemeritCountA { get; set; }
        /// <summary>
        /// 使用者設定 - 小過
        /// </summary>
        public int DemeritCountB { get; set; }
        /// <summary>
        /// 使用者設定 - 警告
        /// </summary>
        public int DemeritCountC { get; set; } 

        #endregion

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
