using K12.BusinessLogic;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    /// <summary>
    /// 獎勵資料處理器
    /// </summary>
    public class TotalObj
    {

        /// <summary>
        /// 懲戒資料清單
        /// 學生ID : 懲戒資料
        /// </summary>
        public Dictionary<string, List<DemeritRecord>> DicByDemerit = new Dictionary<string, List<DemeritRecord>>();

        /// <summary>
        /// 獎勵資料清單
        /// 學生ID : 獎勵資料
        /// </summary>
        public Dictionary<string, List<MeritRecord>> DicByMerit = new Dictionary<string, List<MeritRecord>>();

        /// <summary>
        /// 大於使用者設定之學生ID清單
        /// </summary>
        public List<string> studentUbeIDList = new List<string>();

        /// <summary>
        /// 取得統計資料
        /// </summary>
        public void GetSummary(List<AutoSummaryRecord> AutoSummaryList)
        {
            foreach (AutoSummaryRecord each in AutoSummaryList)
            {
                foreach (MeritRecord merit in each.Merits)
                {
                    if (!DicByMerit.ContainsKey(merit.RefStudentID))
                    {
                        DicByMerit.Add(merit.RefStudentID, new List<MeritRecord>());
                    }
                    DicByMerit[merit.RefStudentID].Add(merit);
                }

                foreach (DemeritRecord demerit in each.Demerits)
                {
                    if (!DicByDemerit.ContainsKey(demerit.RefStudentID))
                    {
                        DicByDemerit.Add(demerit.RefStudentID, new List<DemeritRecord>());
                    }
                    DicByDemerit[demerit.RefStudentID].Add(demerit);
                }
            }
        }

        /// <summary>
        /// 取得依學生編號,日期開始&結束之資料
        /// </summary>
        public void GetDetail(SetConfig _sc)
        {
            List<MeritRecord> meritList = K12.Data.Merit.SelectByOccurDate(_sc.GetStudentIdList(), _sc.StartDate, _sc.EndDate);
            List<DemeritRecord> demeritList = K12.Data.Demerit.SelectByOccurDate(_sc.GetStudentIdList(), _sc.StartDate, _sc.EndDate);

            foreach (MeritRecord merit in meritList)
            {
                if (!DicByMerit.ContainsKey(merit.RefStudentID))
                {
                    DicByMerit.Add(merit.RefStudentID, new List<MeritRecord>());
                }
                DicByMerit[merit.RefStudentID].Add(merit);
            }

            foreach (DemeritRecord demerit in demeritList)
            {
                if (!DicByDemerit.ContainsKey(demerit.RefStudentID))
                {
                    DicByDemerit.Add(demerit.RefStudentID, new List<DemeritRecord>());
                }
                DicByDemerit[demerit.RefStudentID].Add(demerit);
            }
        }
    }
}
