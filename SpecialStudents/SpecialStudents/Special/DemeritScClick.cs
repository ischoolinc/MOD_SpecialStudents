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

        bool _cbxIsMeritAndDemerit;

        SetConfig _sc { get; set; }

        /// <summary>
        /// 資料整理
        /// </summary>
        TotalObj _tb { get; set; }

        public void Print(SetConfig sc, CheckBoxX cbxIsMeritAndDemerit)
        {
            _sc = sc;

            _cbxIsMeritAndDemerit = cbxIsMeritAndDemerit.Checked;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("懲戒特殊表現學生,列印中!");
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();

            AutoSummaryList = new List<AutoSummaryRecord>();

            _sc.GetMerReduce();

            _tb = new TotalObj();

            if (_sc._selectMode == SelectMode.所有學期) //依選擇學期
            {
                //取得AutoSummary
                AutoSummaryList = AutoSummary.Select(_sc.GetStudentIdList(), new List<SchoolYearSemester>(), SummaryType.Discipline);

                _tb.GetSummary(AutoSummaryList);
                _tb.GetInitialSummary(AutoSummaryList);
            }
            else if (_sc._selectMode == SelectMode.依學期)
            {
                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_sc._SchoolYear, _sc._Semester);
                AutoSummaryList = AutoSummary.Select(_sc.GetStudentIdList(), new SchoolYearSemester[] { SYS }, SummaryType.Discipline);

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
            sheet.Name = "懲戒特殊表現學生";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "\n懲戒特殊表現學生";

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

            Aspose.Cells.Row row = sheet.Cells.Rows[0];
            row.Height = 30;

            //A1.PutValue(A1Name);
            obj.FormatCell_2(A1, A1Name);



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

            #endregion

            int index = 1;

            //處理排序問題
            List<string> StudentIDList = new List<string>();
            if (_sc._selectMode == SelectMode.依日期) //依選擇學期
            {
                foreach (string studentID in _tb.DicByDemerit.Keys)
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

                if (_sc._selectMode == SelectMode.依日期) //依選擇學期
                {

                    if (_tb.DicByMerit.ContainsKey(student.ID))
                    {
                        foreach (MeritRecord each in _tb.DicByMerit[student.ID])
                        {
                            int A = each.MeritA.HasValue ? each.MeritA.Value : 0;
                            int B = each.MeritB.HasValue ? each.MeritB.Value : 0;
                            int C = each.MeritC.HasValue ? each.MeritC.Value : 0;

                            //將統計相換算成比值的基底
                            MeritTotal += (A * _sc.Meritab * _sc.Meritbc) + (B * _sc.Meritbc) + (C);


                            MeritA += A;
                            MeritB += B;
                            MeritC += C;
                        }
                    }

                    if (_tb.DicByDemerit.ContainsKey(student.ID))
                    {
                        foreach (DemeritRecord each in _tb.DicByDemerit[student.ID])
                        {
                            if (each.Cleared == "是")
                                continue;

                            int A = each.DemeritA.HasValue ? each.DemeritA.Value : 0;
                            int B = each.DemeritB.HasValue ? each.DemeritB.Value : 0;
                            int C = each.DemeritC.HasValue ? each.DemeritC.Value : 0;

                            DemeritTotal += (A * _sc.Demeritab * _sc.Demeritbc) + (B * _sc.Demeritbc) + (C);

                            DemeritA += A;
                            DemeritB += B;
                            DemeritC += C;

                        }
                    }
                }
                else
                {
                    if (_tb.DicByInitialSummary.ContainsKey(student.ID))
                    {
                        foreach (AutoSummaryRecord each in _tb.DicByInitialSummary[student.ID])
                        {
                            int A = each.MeritA;
                            int B = each.MeritB;
                            int C = each.MeritC;

                            //將統計相換算成比值的基底
                            MeritTotal += (A * _sc.Meritab * _sc.Meritbc) + (B * _sc.Meritbc) + (C);


                            MeritA += A;
                            MeritB += B;
                            MeritC += C;

                            int xA = each.DemeritA;
                            int xB = each.DemeritB;
                            int xC = each.DemeritC;

                            DemeritTotal += (xA * _sc.Demeritab * _sc.Demeritbc) + (xB * _sc.Demeritbc) + (xC);

                            DemeritA += xA;
                            DemeritB += xB;
                            DemeritC += xC;
                        }
                    }
                }



                if (_cbxIsMeritAndDemerit) //進行功過相抵
                {
                    //相減
                    DemeritTotal -= MeritTotal;
                }

                if (DemeritTotal < _sc.Demeritwant) continue; //如果小於基底數,就下一個學生

                _tb.studentUbeIDList.Add(student.ID);

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
            //Cell titleCell = sheet2.Cells["A1"];
            //titleCell.PutValue(School.ChineseName + "\n懲戒明細");
            //titleCell = tool.UserStyle(titleCell);

            //New
            Aspose.Cells.Row row_sheet2 = sheet2.Cells.Rows[0];
            row_sheet2.Height = 30;
            obj.FormatCell_2(sheet2.Cells["A1"], School.ChineseName + "\n懲戒明細");

            sheet2.Cells.Merge(0, 0, 1, 12);

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

            int ri = 3;

            foreach (StudentRecord student in StudentList)
            {
                if (!_tb.studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單,就不印明細
                    continue;

                if (_sc._selectMode == SelectMode.依日期)
                {
                    if (_tb.DicByDemerit.ContainsKey(student.ID))
                    {
                        foreach (DemeritRecord demerit in _tb.DicByDemerit[student.ID])
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
                else
                {
                    if (_tb.DicByDemerit.ContainsKey(student.ID))
                    {
                        foreach (DemeritRecord demerit in _tb.DicByDemerit[student.ID])
                        {
                            if (demerit.Cleared == "是")
                                continue;

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

                    //2021/9/22 - 如果是依學期就使用非明細統計
                    if (_tb.DicByInitialSummary.ContainsKey(student.ID))
                    {
                        foreach (AutoSummaryRecord demerit in _tb.DicByInitialSummary[student.ID])
                        {
                            try
                            {
                                if (demerit.InitialDemeritA + demerit.InitialDemeritB + demerit.InitialDemeritC > 0)
                                {
                                    obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                                    obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                                    obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                                    obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                                    obj.FormatCell(sheet2.Cells["E" + ri], demerit.SchoolYear.ToString());
                                    obj.FormatCell(sheet2.Cells["F" + ri], demerit.Semester.ToString());
                                    obj.FormatCell(sheet2.Cells["G" + ri], "");
                                    obj.FormatCell(sheet2.Cells["H" + ri], "" + demerit.InitialDemeritA);
                                    obj.FormatCell(sheet2.Cells["I" + ri], "" + demerit.InitialDemeritB);
                                    obj.FormatCell(sheet2.Cells["J" + ri], "" + demerit.InitialDemeritC);
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

            #endregion

            if (_cbxIsMeritAndDemerit)
            {
                #region 獎勵明細列印

                int sheetIndex_M = book.Worksheets.Add(); //再加一個Sheet
                Worksheet sheet2_M = book.Worksheets[sheetIndex_M];
                sheet2_M.Name = "獎勵明細";

                //Cell titleCell_M = sheet2_M.Cells["A1"];
                //titleCell_M = tool.UserStyle(titleCell_M);
                //titleCell_M.PutValue(School.ChineseName + "\n獎勵明細");

                Aspose.Cells.Row row_sheet3 = sheet2_M.Cells.Rows[0];
                row_sheet3.Height = 30;
                obj.FormatCell_2(sheet2_M.Cells["A1"], School.ChineseName + "\n獎勵明細");

                sheet2_M.Cells.Merge(0, 0, 1, 12);

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

                int ri_M = 3;

                foreach (StudentRecord student in StudentList)
                {
                    if (!_tb.studentUbeIDList.Contains(student.ID)) //如果不包含於列印清單,就不印明細
                        continue;

                    if (_sc._selectMode != SelectMode.所有學期)
                    {
                        if (_tb.DicByMerit.ContainsKey(student.ID))
                        {
                            foreach (MeritRecord merit in _tb.DicByMerit[student.ID])
                            {
                                obj.FormatCell(sheet2_M.Cells["A" + ri_M], student.Class.Name);
                                obj.FormatCell(sheet2_M.Cells["B" + ri_M], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["C" + ri_M], student.Name);
                                obj.FormatCell(sheet2_M.Cells["D" + ri_M], student.StudentNumber);
                                obj.FormatCell(sheet2_M.Cells["E" + ri_M], merit.SchoolYear.ToString());
                                obj.FormatCell(sheet2_M.Cells["F" + ri_M], merit.Semester.ToString());
                                obj.FormatCell(sheet2_M.Cells["G" + ri_M], merit.OccurDate.ToShortDateString());
                                obj.FormatCell(sheet2_M.Cells["H" + ri_M], merit.MeritA.HasValue ? merit.MeritA.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["I" + ri_M], merit.MeritB.HasValue ? merit.MeritB.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["J" + ri_M], merit.MeritC.HasValue ? merit.MeritC.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["K" + ri_M], merit.Reason);
                                obj.FormatCell(sheet2_M.Cells["L" + ri_M], merit.RegisterDate.HasValue ? merit.RegisterDate.Value.ToShortDateString() : "");
                                ri_M++;
                            }
                        }
                    }
                    else
                    {
                        if (_tb.DicByMerit.ContainsKey(student.ID))
                        {
                            foreach (MeritRecord merit in _tb.DicByMerit[student.ID])
                            {
                                obj.FormatCell(sheet2_M.Cells["A" + ri_M], student.Class.Name);
                                obj.FormatCell(sheet2_M.Cells["B" + ri_M], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["C" + ri_M], student.Name);
                                obj.FormatCell(sheet2_M.Cells["D" + ri_M], student.StudentNumber);
                                obj.FormatCell(sheet2_M.Cells["E" + ri_M], merit.SchoolYear.ToString());
                                obj.FormatCell(sheet2_M.Cells["F" + ri_M], merit.Semester.ToString());
                                obj.FormatCell(sheet2_M.Cells["G" + ri_M], merit.OccurDate.ToShortDateString());
                                obj.FormatCell(sheet2_M.Cells["H" + ri_M], merit.MeritA.HasValue ? merit.MeritA.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["I" + ri_M], merit.MeritB.HasValue ? merit.MeritB.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["J" + ri_M], merit.MeritC.HasValue ? merit.MeritC.Value.ToString() : "");
                                obj.FormatCell(sheet2_M.Cells["K" + ri_M], merit.Reason);
                                obj.FormatCell(sheet2_M.Cells["L" + ri_M], merit.RegisterDate.HasValue ? merit.RegisterDate.Value.ToShortDateString() : "");
                                ri_M++;
                            }
                        }

                        //2021/9/22 - 如果是依學期就使用明細統計
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
                #endregion
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SpecialEvent.RaiseSpecialChanged();
            if (!e.Cancelled)
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
            else
            {
                MsgBox.Show("列印作業已中止!");
            }
        }
    }
}
