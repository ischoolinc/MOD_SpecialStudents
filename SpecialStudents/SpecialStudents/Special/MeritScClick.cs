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

        bool _cbxIgnoreDemerit;
        bool _cbxDemeritIsNull;
        bool _cbxIsDemeritClear;

        /// <summary>
        /// 設定內容
        /// </summary>
        SetConfig _sc { get; set; }

        /// <summary>
        /// 資料整理
        /// </summary>
        TotalObj _tb { get; set; }

        public void print(SetConfig sc, CheckBoxX cbxIgnoreDemerit, CheckBoxX cbxDemeritIsNull, CheckBoxX cbxIsDemeritClear)
        {
            _sc = sc;

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

            AutoSummaryList = new List<AutoSummaryRecord>();

            List<string> _StudentIDList = _sc.GetStudentIdList();

            _sc.GetMerReduce();

            _tb = new TotalObj();

            if (_sc._selectMode == SelectMode.所有學期) //依選擇學期
            {
                //取得AutoSummary
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Discipline);
                _tb.GetSummary(AutoSummaryList);
                _tb.GetInitialSummary(AutoSummaryList);
            }
            else if (_sc._selectMode == SelectMode.依學期)
            {
                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_sc._SchoolYear, _sc._Semester);
                AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Discipline);

                _tb.GetSummary(AutoSummaryList);
                _tb.GetInitialSummary(AutoSummaryList);
            }
            else
            {
                _tb.GetDetail(_sc);
            }

            #region 表頭&相關資料準備
            book = new Workbook();
            book.Worksheets.Clear();
            int SHEETIndex = book.Worksheets.Add();
            Worksheet sheet = book.Worksheets[SHEETIndex];
            sheet.Name = "獎勵特殊表現學生";
            //string wantString = wa + " 大功 " + wb + " 小功 " + wc + " 嘉獎";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "\n獎勵特殊表現學生";

            if (_sc._selectMode == SelectMode.依學期)
            {
                A1Name += "　(" + _sc._SchoolYear.ToString() + "/" + _sc._Semester.ToString() + ")";
            }
            else if (_sc._selectMode == SelectMode.依日期)
            {
                A1Name += "　(" + _sc.StartDate.ToShortDateString() + "~" + _sc.EndDate.ToShortDateString() + ")";
            }
            else
            {
                A1Name += "(所有學期)";
            }

            //obj.FormatCell(A1, A1Name);
            //A1.PutValue(A1Name);

            Aspose.Cells.Row row = sheet.Cells.Rows[0];
            row.Height = 30;
            obj.FormatCell_2(A1, A1Name);

            sheet.Cells.Merge(0, 0, 1, 7);

            sheet.Cells.Columns[0].Width = 10;
            sheet.Cells.Columns[1].Width = 10;
            sheet.Cells.Columns[2].Width = 10;
            sheet.Cells.Columns[3].Width = 10;
            sheet.Cells.Columns[4].Width = 10;
            sheet.Cells.Columns[5].Width = 10;
            sheet.Cells.Columns[6].Width = 10;
            sheet.Cells.Columns[7].Width = 10;

            obj.FormatCell(sheet.Cells["A2"], "班級");
            obj.FormatCell(sheet.Cells["B2"], "座號");
            obj.FormatCell(sheet.Cells["C2"], "姓名");
            obj.FormatCell(sheet.Cells["D2"], "學號");
            obj.FormatCell(sheet.Cells["E2"], "大功");
            obj.FormatCell(sheet.Cells["F2"], "小功");
            obj.FormatCell(sheet.Cells["G2"], "嘉獎");

            //AutoSummaryList.Sort(new SortClass().SortAutoSummaryRecord); 

            #endregion

            int index = 1;

            #region 處理獎勵資料列印1

            //處理排序問題
            List<string> StudentIDList = new List<string>();

            if (_sc._selectMode == SelectMode.依日期) //依選擇學期
            {
                foreach (string studentID in _tb.DicByMerit.Keys)
                {
                    if (!StudentIDList.Contains(studentID))
                    {
                        StudentIDList.Add(studentID);
                    }
                }
            }
            else
            {
                foreach (string studentID in _tb.DicByInitialSummary.Keys)
                {
                    if (!StudentIDList.Contains(studentID))
                    {
                        StudentIDList.Add(studentID);
                    }
                }

            }

            List<StudentRecord> StudentList = SortClassIndex.K12Data_StudentRecord(Student.SelectByIDs(StudentIDList));


            foreach (StudentRecord student in StudentList)
            {
                if (UsingDemeritData(student.ID, _cbxIgnoreDemerit, _cbxDemeritIsNull, _cbxIsDemeritClear)) //依使用者選擇的條件進行處理
                    continue;

                //將統計相換算成比值的基底
                int total = 0;
                int MeritA = 0;
                int MeritB = 0;
                int MeritC = 0;

                if (_sc._selectMode == SelectMode.依日期)
                {
                    if (_tb.DicByMerit.ContainsKey(student.ID))
                    {
                        foreach (MeritRecord each in _tb.DicByMerit[student.ID])
                        {
                            int A = each.MeritA.HasValue ? each.MeritA.Value : 0;
                            int B = each.MeritB.HasValue ? each.MeritB.Value : 0;
                            int C = each.MeritC.HasValue ? each.MeritC.Value : 0;

                            total += (A * _sc.Meritab * _sc.Meritbc) + (B * _sc.Meritbc) + (C);

                            MeritA += A;
                            MeritB += B;
                            MeritC += C;
                        }
                    }
                }
                else
                {
                    if (_tb.DicByInitialSummary.ContainsKey(student.ID))
                    {
                        foreach (AutoSummaryRecord summary in _tb.DicByInitialSummary[student.ID])
                        {
                            int A = summary.MeritA;
                            int B = summary.MeritB;
                            int C = summary.MeritC;

                            total += (A * _sc.Meritab * _sc.Meritbc) + (B * _sc.Meritbc) + (C);

                            MeritA += A;
                            MeritB += B;
                            MeritC += C;

                        }
                    }
                }

                //如果小於基底數,就下一個學生
                if (total < _sc.Meritwant) continue;

                _tb.studentUbeIDList.Add(student.ID);

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

            //Cell titleCell = sheet2.Cells["A1"];
            //titleCell = tool.UserStyle(titleCell);
            //titleCell.PutValue(School.ChineseName + "\n獎勵累計明細");

            Aspose.Cells.Row row_sheet2 = sheet2.Cells.Rows[0];
            row_sheet2.Height = 30;
            obj.FormatCell_2(sheet2.Cells["A1"], School.ChineseName + "\n獎勵累計明細");

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

                if (!_tb.studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單
                    continue;

                if (_sc._selectMode == SelectMode.依日期)
                {
                    if (_tb.DicByMerit.ContainsKey(student.ID))
                    {
                        foreach (MeritRecord merit in _tb.DicByMerit[student.ID])
                        {
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
                else
                {
                    if (_tb.DicByMerit.ContainsKey(student.ID))
                    {
                        foreach (MeritRecord merit in _tb.DicByMerit[student.ID])
                        {
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

                    //2021/9/22 - 如果是依學期就使用非明細統計
                    if (_tb.DicByInitialSummary.ContainsKey(student.ID))
                    {
                        foreach (AutoSummaryRecord summary in _tb.DicByInitialSummary[student.ID])
                        {
                            try
                            {
                                if (summary.InitialMeritA + summary.InitialMeritB + summary.InitialMeritC > 0)
                                {
                                    obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                                    obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                                    obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                                    obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                                    obj.FormatCell(sheet2.Cells["E" + ri], "" + summary.SchoolYear);
                                    obj.FormatCell(sheet2.Cells["F" + ri], "" + summary.Semester);
                                    obj.FormatCell(sheet2.Cells["G" + ri], "");
                                    obj.FormatCell(sheet2.Cells["H" + ri], "" + summary.InitialMeritA);
                                    obj.FormatCell(sheet2.Cells["I" + ri], "" + summary.InitialMeritB);
                                    obj.FormatCell(sheet2.Cells["J" + ri], "" + summary.InitialMeritC);
                                    obj.FormatCell(sheet2.Cells["K" + ri], "(非明細資料)");
                                    obj.FormatCell(sheet2.Cells["L" + ri], "");

                                    ri++;
                                }
                            }
                            catch
                            {
                                //沒有 Initial 就把錯誤吃掉...
                            }
                        }
                    }
                }
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SpecialEvent.RaiseSpecialChanged();

            if (!e.Cancelled)
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
            else
            {
                MsgBox.Show("列印作業已中止!");
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
                if (_tb.DicByDemerit.ContainsKey(StudentID)) //如果有此學生,表示有懲戒記錄
                {
                    return true;
                }
            }
            else if (cbxIsDemeritClear) //有懲戒記錄,都已銷過者才列入清單
            {
                if (_tb.DicByDemerit.ContainsKey(StudentID)) //如果有此學生,表示有懲戒記錄
                {
                    bool CheckDemerit = false;

                    foreach (DemeritRecord demeirt in _tb.DicByDemerit[StudentID])
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
