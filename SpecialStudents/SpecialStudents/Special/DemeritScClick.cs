using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using FISCA.Presentation.Controls;
using DevComponents.DotNetBar.Controls;
using DevComponents.Editors;
using K12.Data;
using FISCA.DSAUtil;
using System.Drawing;
using System.ComponentModel;
using K12.BusinessLogic;

namespace SpecialStudents
{
    class DemeritScClick
    {
        PrintObj obj;

        BackgroundWorker BGW = new BackgroundWorker();

        //報表專用
        Workbook book;

        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        string _tbDemeritA;
        string _tbDemeritB;
        string _tbDemeritC;
        bool _cbxIsMeritAndDemerit;

        SetConfig _sc { get; set; }

        public void Print(SetConfig sc, TextBoxX tbDemeritA, TextBoxX tbDemeritB, TextBoxX tbDemeritC, CheckBoxX cbxIsMeritAndDemerit)
        {
            _sc = sc;

            _tbDemeritA = tbDemeritA.Text;
            _tbDemeritB = tbDemeritB.Text;
            _tbDemeritC = tbDemeritC.Text;
            _cbxIsMeritAndDemerit = cbxIsMeritAndDemerit.Checked;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("懲戒特殊表現學生,列印中!");
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();
            if (obj.CheckTextBox(_tbDemeritA, _tbDemeritB, _tbDemeritC))
            {
                MsgBox.Show("懲戒次數必須輸入數字!");
                return;
            }

            AutoSummaryList = new List<AutoSummaryRecord>();
            List<string> _StudentIDList = _sc._StudentList.Select(x => x.ID).ToList();

            List<string> studentUbeIDList = new List<string>(); //被列印的學生,用以印明細時的判斷

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
            MeritDemeritReduceRecord MDRecord = MeritDemeritReduce.Select();

            //過Demerit
            int Demeritab = MDRecord.DemeritAToDemeritB.HasValue ? MDRecord.DemeritAToDemeritB.Value : 1;
            int Demeritbc = MDRecord.DemeritBToDemeritC.HasValue ? MDRecord.DemeritBToDemeritC.Value : 1;

            //獎Demerit
            int Meritab = MDRecord.MeritAToMeritB.HasValue ? MDRecord.MeritAToMeritB.Value : 1;
            int Meritbc = MDRecord.MeritBToMeritC.HasValue ? MDRecord.MeritBToMeritC.Value : 1;

            //三個欄位的內容
            //大過
            int wa = int.Parse(_tbDemeritA);
            //小過
            int wb = int.Parse(_tbDemeritB);
            //警告
            int wc = int.Parse(_tbDemeritC);

            //計算出基數
            int want = (wa * Demeritab * Demeritbc) + (wb * Demeritbc) + wc;

            #endregion

            #region 表頭&相關資料準備
            book = new Workbook();
            book.Worksheets.Clear();
            int SHEETIndex = book.Worksheets.Add();
            Worksheet sheet = book.Worksheets[SHEETIndex];
            sheet.Name = "懲戒特殊表現學生";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "　懲戒特殊表現學生";
            if (_sc._selectMode == SelectMode.依學期)
            {
                A1Name += "　(" + _sc._SchoolYear + "/" + _sc._Semester + ")";
            }

            A1.PutValue(A1Name);

            if (_cbxIsMeritAndDemerit)
            {
                sheet.Cells.Merge(0, 0, 1, 11);
            }
            else
            {
                sheet.Cells.Merge(0, 0, 1, 8);
            }

            obj.FormatCell(sheet.Cells["A2"], "班級");
            obj.FormatCell(sheet.Cells["B2"], "座號");
            obj.FormatCell(sheet.Cells["C2"], "姓名");
            obj.FormatCell(sheet.Cells["D2"], "學號");
            obj.FormatCell(sheet.Cells["E2"], "大過");
            obj.FormatCell(sheet.Cells["F2"], "小過");
            obj.FormatCell(sheet.Cells["G2"], "警告");
            if (_cbxIsMeritAndDemerit)
            {
                obj.FormatCell(sheet.Cells["H2"], "大功");
                obj.FormatCell(sheet.Cells["I2"], "小功");
                obj.FormatCell(sheet.Cells["J2"], "嘉獎");
                obj.FormatCell(sheet.Cells["K2"], "單位(次)");
            }
            else
            {
                obj.FormatCell(sheet.Cells["H2"], "單位(次)");
            }

            studentUbeIDList.Clear();
            //AutoSummaryList.Sort(new SortClass().SortAutoSummaryRecord); 
            #endregion

            int index = 1;

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

            #region 列印資料1

            foreach (StudentRecord student in StudentList)
            {
                int DemeritTotal = 0;
                int MeritTotal = 0;

                int DemeritA = 0;
                int DemeritB = 0;
                int DemeritC = 0;
                int MeritA = 0;
                int MeritB = 0;
                int MeritC = 0;

                foreach (AutoSummaryRecord each in AutoList[student.ID])
                {
                    //將統計相換算成比值的基底
                    DemeritTotal += (each.DemeritA * Demeritab * Demeritbc) + (each.DemeritB * Demeritbc) + (each.DemeritC);
                    MeritTotal += (each.MeritA * Meritab * Meritbc) + (each.MeritB * Meritbc) + (each.MeritC);
                    DemeritA += each.DemeritA;
                    DemeritB += each.DemeritB;
                    DemeritC += each.DemeritC;
                    MeritA += each.MeritA;
                    MeritB += each.MeritB;
                    MeritC += each.MeritC;
                }

                if (_cbxIsMeritAndDemerit) //進行功過相抵
                {
                    //相減
                    DemeritTotal -= MeritTotal;
                }

                if (DemeritTotal < want || DemeritTotal == 0) continue; //如果小於基底數,就下一個學生

                studentUbeIDList.Add(student.ID);

                int rowIndex = index + 2;
                obj.FormatCell(sheet.Cells["A" + rowIndex], student.Class.Name);
                obj.FormatCell(sheet.Cells["B" + rowIndex], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                obj.FormatCell(sheet.Cells["C" + rowIndex], student.Name);
                obj.FormatCell(sheet.Cells["D" + rowIndex], student.StudentNumber);
                obj.FormatCell(sheet.Cells["E" + rowIndex], DemeritA.ToString()); //大過
                obj.FormatCell(sheet.Cells["F" + rowIndex], DemeritB.ToString()); //小過
                obj.FormatCell(sheet.Cells["G" + rowIndex], DemeritC.ToString()); //警告
                if (_cbxIsMeritAndDemerit)
                {
                    obj.FormatCell(sheet.Cells["H" + rowIndex], MeritA.ToString()); //大功
                    obj.FormatCell(sheet.Cells["I" + rowIndex], MeritB.ToString()); //小功
                    obj.FormatCell(sheet.Cells["J" + rowIndex], MeritC.ToString()); //嘉獎
                    obj.FormatCell(sheet.Cells["K" + rowIndex], DemeritTotal.ToString()); //單位次

                }
                else
                {
                    obj.FormatCell(sheet.Cells["H" + rowIndex], DemeritTotal.ToString()); //單位次
                }
                index++;
            }
            #endregion

            #region 懲戒明細列印

            int sheetIndex = book.Worksheets.Add(); //再加一個Sheet
            Worksheet sheet2 = book.Worksheets[sheetIndex];
            sheet2.Name = "懲戒明細";
            Cell titleCell = sheet2.Cells["A1"];
            titleCell.PutValue(School.ChineseName + "　懲戒明細");

            titleCell = tool.UserStyle(titleCell);

            sheet2.Cells.Merge(0, 0, 1, 12);

            #region 欄位Title
            obj.FormatCell(sheet2.Cells["A2"], "班級");
            obj.FormatCell(sheet2.Cells["B2"], "座號");
            obj.FormatCell(sheet2.Cells["C2"], "姓名");
            obj.FormatCell(sheet2.Cells["D2"], "學號");
            obj.FormatCell(sheet2.Cells["E2"], "學年度");
            obj.FormatCell(sheet2.Cells["F2"], "學期");
            obj.FormatCell(sheet2.Cells["G2"], "發生日期");
            obj.FormatCell(sheet2.Cells["H2"], "大過");
            obj.FormatCell(sheet2.Cells["I2"], "小過");
            obj.FormatCell(sheet2.Cells["J2"], "警告");
            //obj.FormatCell(sheet2.Cells["K2"], "留察"); //留察在國中系統屬於特別懲戒...
            obj.FormatCell(sheet2.Cells["K2"], "事由");
            obj.FormatCell(sheet2.Cells["L2"], "登錄日期");

            //obj.FormatCell(sheet2.Cells["L2"], "是否銷過");
            //obj.FormatCell(sheet2.Cells["M2"], "銷過日期");
            //obj.FormatCell(sheet2.Cells["N2"], "銷過事由");
            #endregion

            #region 列印資料2

            int ri = 3;

            foreach (StudentRecord student in StudentList)
            {
                if (!studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單,就不印明細
                    continue;

                foreach (AutoSummaryRecord each in AutoList[student.ID])
                {
                    foreach (DemeritRecord demerit in each.Demerits)
                    {
                        if (demerit.Cleared == "是")
                            continue;

                        //StudentRecord student = JHStudent.SelectByID(demerit.RefStudentID); //取得學生

                        obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                        obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                        obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                        obj.FormatCell(sheet2.Cells["E" + ri], demerit.SchoolYear.ToString());
                        obj.FormatCell(sheet2.Cells["F" + ri], demerit.Semester.ToString());
                        obj.FormatCell(sheet2.Cells["G" + ri], demerit.OccurDate.ToShortDateString());
                        obj.FormatCell(sheet2.Cells["H" + ri], demerit.DemeritA.HasValue ? demerit.DemeritA.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["I" + ri], demerit.DemeritB.HasValue ? demerit.DemeritB.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["J" + ri], demerit.DemeritC.HasValue ? demerit.DemeritC.Value.ToString() : "");
                        obj.FormatCell(sheet2.Cells["K" + ri], demerit.Reason);
                        obj.FormatCell(sheet2.Cells["L" + ri], demerit.RegisterDate.HasValue ? demerit.RegisterDate.Value.ToShortDateString() : "");
                        ri++;
                    }
                }
            }
            #endregion
            #endregion

            if (_cbxIsMeritAndDemerit)
            {
                #region 獎勵明細列印

                int sheetIndex_M = book.Worksheets.Add(); //再加一個Sheet
                Worksheet sheet2_M = book.Worksheets[sheetIndex_M];
                sheet2_M.Name = "獎勵明細";
                Cell titleCell_M = sheet2_M.Cells["A1"];
                titleCell_M = tool.UserStyle(titleCell_M);

                titleCell_M.PutValue(School.ChineseName + "　獎勵明細");

                sheet2_M.Cells.Merge(0, 0, 1, 12);

                #region 欄位Title
                obj.FormatCell(sheet2_M.Cells["A2"], "班級");
                obj.FormatCell(sheet2_M.Cells["B2"], "座號");
                obj.FormatCell(sheet2_M.Cells["C2"], "姓名");
                obj.FormatCell(sheet2_M.Cells["D2"], "學號");
                obj.FormatCell(sheet2_M.Cells["E2"], "學年度");
                obj.FormatCell(sheet2_M.Cells["F2"], "學期");
                obj.FormatCell(sheet2_M.Cells["G2"], "發生日期");
                obj.FormatCell(sheet2_M.Cells["H2"], "大功");
                obj.FormatCell(sheet2_M.Cells["I2"], "小功");
                obj.FormatCell(sheet2_M.Cells["J2"], "嘉獎");
                obj.FormatCell(sheet2_M.Cells["K2"], "事由");
                obj.FormatCell(sheet2_M.Cells["L2"], "登錄日期");
                #endregion

                #region 列印資料3

                int ri_M = 3;

                foreach (StudentRecord student in StudentList)
                {
                    if (!studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單,就不印明細
                        continue;

                    foreach (AutoSummaryRecord each in AutoList[student.ID])
                    {
                        foreach (MeritRecord Merit in each.Merits)
                        {
                            obj.FormatCell(sheet2_M.Cells["A" + ri_M], student.Class.Name);
                            obj.FormatCell(sheet2_M.Cells["B" + ri_M], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                            obj.FormatCell(sheet2_M.Cells["C" + ri_M], student.Name);
                            obj.FormatCell(sheet2_M.Cells["D" + ri_M], student.StudentNumber);
                            obj.FormatCell(sheet2_M.Cells["E" + ri_M], Merit.SchoolYear.ToString());
                            obj.FormatCell(sheet2_M.Cells["F" + ri_M], Merit.Semester.ToString());
                            obj.FormatCell(sheet2_M.Cells["G" + ri_M], Merit.OccurDate.ToShortDateString());
                            obj.FormatCell(sheet2_M.Cells["H" + ri_M], Merit.MeritA.HasValue ? Merit.MeritA.Value.ToString() : "");
                            obj.FormatCell(sheet2_M.Cells["I" + ri_M], Merit.MeritB.HasValue ? Merit.MeritB.Value.ToString() : "");
                            obj.FormatCell(sheet2_M.Cells["J" + ri_M], Merit.MeritC.HasValue ? Merit.MeritC.Value.ToString() : "");
                            obj.FormatCell(sheet2_M.Cells["K" + ri_M], Merit.Reason);
                            obj.FormatCell(sheet2_M.Cells["L" + ri_M], Merit.RegisterDate.HasValue ? Merit.RegisterDate.Value.ToShortDateString() : "");
                            ri_M++;
                        }
                    }
                }
                #endregion

                #endregion
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                obj.PrintNow(book, "懲戒特殊表現學生");
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印懲戒特殊表現學生,已完成!");
            }
            else
            {
                MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage("列印懲戒特殊表現學生,發生錯誤!");
            }
        }
    }
}
