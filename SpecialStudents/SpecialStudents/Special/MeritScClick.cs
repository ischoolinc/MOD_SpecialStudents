using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using DevComponents.Editors;
using DevComponents.DotNetBar.Controls;
using K12.Data;
using FISCA.DSAUtil;
using System.Drawing;
using System.ComponentModel;
using K12.BusinessLogic;

namespace SpecialStudents
{
    class MeritScClick
    {
        PrintObj obj;

        BackgroundWorker BGW = new BackgroundWorker();

        //報表專用
        Workbook book;

        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        //懲戒資料清單
        Dictionary<string, List<DemeritRecord>> DicByDemerit = new Dictionary<string, List<DemeritRecord>>();
        //獎勵資料清單
        Dictionary<string, List<MeritRecord>> DicByMerit = new Dictionary<string, List<MeritRecord>>();

        string _tbMeritA;
        string _tbMeritB;
        string _tbMeritC;
        bool _cbxIgnoreDemerit;
        bool _cbxDemeritIsNull;
        bool _cbxIsDemeritClear;

        SetConfig _sc { get; set; }

        public void print(SetConfig sc, TextBoxX tbMeritA, TextBoxX tbMeritB, TextBoxX tbMeritC,
            CheckBoxX cbxIgnoreDemerit, CheckBoxX cbxDemeritIsNull, CheckBoxX cbxIsDemeritClear)
        {
            _sc = sc;

            _tbMeritA = tbMeritA.Text;
            _tbMeritB = tbMeritB.Text;
            _tbMeritC = tbMeritC.Text;
            _cbxIgnoreDemerit = cbxIgnoreDemerit.Checked;
            _cbxDemeritIsNull = cbxDemeritIsNull.Checked;
            _cbxIsDemeritClear = cbxIsDemeritClear.Checked;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("獎勵特殊表現學生,列印中!");
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();
            if (obj.CheckTextBox(_tbMeritA, _tbMeritB, _tbMeritC))
            {
                MsgBox.Show("獎勵次數必須輸入數字!");
                return;
            }

            AutoSummaryList = new List<AutoSummaryRecord>();

            List<string> _StudentIDList = _sc._StudentList.Select(x => x.ID).ToList();
            List<string> studentUbeIDList = new List<string>();

            if (_sc._selectMode == SelectMode.所有學期) //依選擇學期
            {
                //取得AutoSummary
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Discipline);

            }
            else if (_sc._selectMode == SelectMode.依日期)
            {




            }
            else
            {
                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_sc._SchoolYear, _sc._Semester);
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Discipline);
            }

            #region 取得換算原則

            MeritDemeritReduceRecord record = K12.Data.MeritDemeritReduce.Select();
            int ab = record.MeritAToMeritB.HasValue ? record.MeritAToMeritB.Value : 1;
            int bc = record.MeritBToMeritC.HasValue ? record.MeritBToMeritC.Value : 1;

            //三個欄位的內容
            //大功
            int wa = int.Parse(_tbMeritA);
            //小功
            int wb = int.Parse(_tbMeritB);
            //嘉獎
            int wc = int.Parse(_tbMeritC);

            //計算出基數
            int want = (wa * ab * bc) + (wb * bc) + wc;
            #endregion

            #region 表頭&相關資料準備
            book = new Workbook();
            book.Worksheets.Clear();
            int SHEETIndex = book.Worksheets.Add();
            Worksheet sheet = book.Worksheets[SHEETIndex];
            sheet.Name = "獎勵特殊表現學生";
            //string wantString = wa + " 大功 " + wb + " 小功 " + wc + " 嘉獎";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "　獎勵特殊表現學生";

            if (_sc._selectMode == SelectMode.依學期)
            {
                A1Name += "　(" + _sc._SchoolYear.ToString() + "/" + _sc._Semester.ToString() + ")";
            }

            A1.PutValue(A1Name);

            sheet.Cells.Merge(0, 0, 1, 7);

            obj.FormatCell(sheet.Cells["A2"], "班級");
            obj.FormatCell(sheet.Cells["B2"], "座號");
            obj.FormatCell(sheet.Cells["C2"], "姓名");
            obj.FormatCell(sheet.Cells["D2"], "學號");
            obj.FormatCell(sheet.Cells["E2"], "大功");
            obj.FormatCell(sheet.Cells["F2"], "小功");
            obj.FormatCell(sheet.Cells["G2"], "嘉獎");

            DicByDemerit.Clear();
            DicByMerit.Clear();
            studentUbeIDList.Clear();

            //AutoSummaryList.Sort(new SortClass().SortAutoSummaryRecord); 

            #endregion

            #region 整理獎懲資料清單
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
            #endregion

            int index = 1;

            #region 處理獎勵資料列印1

            //處理排序問題
            Dictionary<string, List<AutoSummaryRecord>> AutoList = new Dictionary<string, List<AutoSummaryRecord>>();
            foreach (AutoSummaryRecord each in AutoSummaryList)
            {
                if (!AutoList.ContainsKey(each.RefStudentID))
                {
                    AutoList.Add(each.RefStudentID, new List<AutoSummaryRecord>());
                }

                AutoList[each.RefStudentID].Add(each);
            }
            List<StudentRecord> StudentList = Student.SelectByIDs(AutoList.Keys);
            StudentList = SortClassIndex.K12Data_StudentRecord(StudentList);


            foreach (StudentRecord student in StudentList)
            {
                if (UsingDemeritData(student.ID, _cbxIgnoreDemerit, _cbxDemeritIsNull, _cbxIsDemeritClear)) //依使用者選擇的條件進行處理
                    continue;

                //將統計相換算成比值的基底
                int total = 0;
                int MeritA = 0;
                int MeritB = 0;
                int MeritC = 0;

                foreach (AutoSummaryRecord each in AutoList[student.ID])
                {
                    total += (each.MeritA * ab * bc) + (each.MeritB * bc) + (each.MeritC);

                    MeritA += each.MeritA;
                    MeritB += each.MeritB;
                    MeritC += each.MeritC;
                }

                if (total < want || total == 0) continue; //如果小於基底數,就下一個學生

                studentUbeIDList.Add(student.ID);

                int rowIndex = index + 2;
                obj.FormatCell(sheet.Cells["A" + rowIndex], student.Class.Name);
                obj.FormatCell(sheet.Cells["B" + rowIndex], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                obj.FormatCell(sheet.Cells["C" + rowIndex], student.Name);
                obj.FormatCell(sheet.Cells["D" + rowIndex], student.StudentNumber);
                obj.FormatCell(sheet.Cells["E" + rowIndex], MeritA.ToString());
                obj.FormatCell(sheet.Cells["F" + rowIndex], MeritB.ToString());
                obj.FormatCell(sheet.Cells["G" + rowIndex], MeritC.ToString());
                index++;
            }
            #endregion

            int sheetIndex = book.Worksheets.Add(); //再加一個Sheet
            Worksheet sheet2 = book.Worksheets[sheetIndex];
            sheet2.Name = "獎勵累計明細";

            Cell titleCell = sheet2.Cells["A1"];

            titleCell = tool.UserStyle(titleCell);

            titleCell.PutValue(School.ChineseName + "　獎勵累計明細");
            sheet2.Cells.Merge(0, 0, 1, 12);

            #region 欄位Title
            obj.FormatCell(sheet2.Cells["A2"], "班級");
            obj.FormatCell(sheet2.Cells["B2"], "座號");
            obj.FormatCell(sheet2.Cells["C2"], "姓名");
            obj.FormatCell(sheet2.Cells["D2"], "學號");
            obj.FormatCell(sheet2.Cells["E2"], "學年度");
            obj.FormatCell(sheet2.Cells["F2"], "學期");
            obj.FormatCell(sheet2.Cells["G2"], "發生日期");
            obj.FormatCell(sheet2.Cells["H2"], "大功");
            obj.FormatCell(sheet2.Cells["I2"], "小功");
            obj.FormatCell(sheet2.Cells["J2"], "嘉獎");
            obj.FormatCell(sheet2.Cells["K2"], "事由");
            obj.FormatCell(sheet2.Cells["L2"], "登錄日期");
            #endregion

            int ri = 3;

            foreach (StudentRecord student in StudentList)
            {
                if (UsingDemeritData(student.ID, _cbxIgnoreDemerit, _cbxDemeritIsNull, _cbxIsDemeritClear))//依使用者選擇的條件進行處理
                    continue;

                if (!studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單
                    continue;

                foreach (AutoSummaryRecord each in AutoList[student.ID])
                {
                    foreach (MeritRecord merit in each.Merits)
                    {

                        //StudentRecord student = JHStudent.SelectByID(each.RefStudentID); //取得學生

                        obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                        obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                        obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                        obj.FormatCell(sheet2.Cells["E" + ri], merit.SchoolYear.ToString());
                        obj.FormatCell(sheet2.Cells["F" + ri], merit.Semester.ToString());
                        obj.FormatCell(sheet2.Cells["G" + ri], merit.OccurDate.ToShortDateString());
                        obj.FormatCell(sheet2.Cells["H" + ri], merit.MeritA.HasValue ? merit.MeritA.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["I" + ri], merit.MeritB.HasValue ? merit.MeritB.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["J" + ri], merit.MeritC.HasValue ? merit.MeritC.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["K" + ri], merit.Reason);
                        obj.FormatCell(sheet2.Cells["L" + ri], merit.RegisterDate.HasValue ? merit.RegisterDate.Value.ToShortDateString() : "");

                        ri++;
                    }
                }
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                obj.PrintNow(book, "獎勵特殊表現學生");
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印獎勵特殊表現學生,已完成!");
            }
            else
            {
                MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印獎勵特殊表現學生,發生錯誤!");
            }
        }

        //依條件判斷是否列印獎勵資料
        private bool UsingDemeritData(string StudentID, bool cbxIgnoreDemerit, bool cbxDemeritIsNull, bool cbxIsDemeritClear)
        {
            if (cbxIgnoreDemerit) //忽略懲戒記錄
            {
                return false;
                //無條件繼續執行
            }
            else if (cbxDemeritIsNull) //有懲戒記錄者,不列入清單
            {
                if (DicByDemerit.ContainsKey(StudentID)) //如果有此學生,表示有懲戒記錄
                {
                    return true;
                }
            }
            else if (cbxIsDemeritClear) //有懲戒記錄,都已銷過者才列入清單
            {
                if (DicByDemerit.ContainsKey(StudentID)) //如果有此學生,表示有懲戒記錄
                {
                    bool CheckDemerit = false;

                    foreach (DemeritRecord demeirt in DicByDemerit[StudentID])
                    {
                        if (demeirt.Cleared != "是")
                        {
                            CheckDemerit = true; //有資料未銷過
                        }
                    }

                    if (CheckDemerit) //如果有資料未銷過,則略過
                        return true;
                }
            }

            return false;

        }
    }
}
