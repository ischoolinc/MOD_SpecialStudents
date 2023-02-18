using Aspose.Cells;
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;
using K12.BusinessLogic;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    internal class MerDemerAllScClick
    {
        PrintObj obj;

        BackgroundWorker BGW = new BackgroundWorker();

        //報表專用
        Workbook book;

        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        /// <summary>
        /// 設定內容
        /// </summary>
        SetConfig _sc { get; set; }

        /// <summary>
        /// 資料整理
        /// </summary>
        TotalObj _tb { get; set; }

        public void Print(SetConfig sc)
        {
            _sc = sc;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("獎懲總計,列印中!");
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();

            AutoSummaryList = new List<AutoSummaryRecord>();

            List<string> _StudentIDList = _sc.GetStudentIdList();


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
            sheet.Name = "獎懲總計";
            //string wantString = wa + " 大功 " + wb + " 小功 " + wc + " 嘉獎";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "\n獎懲總計";

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

            sheet.Cells.Merge(0, 0, 1, 10);

            sheet.Cells.Columns[0].Width = 10;
            sheet.Cells.Columns[1].Width = 10;
            sheet.Cells.Columns[2].Width = 10;
            sheet.Cells.Columns[3].Width = 10;
            sheet.Cells.Columns[4].Width = 10;
            sheet.Cells.Columns[5].Width = 10;
            sheet.Cells.Columns[6].Width = 10;
            sheet.Cells.Columns[7].Width = 10;
            sheet.Cells.Columns[8].Width = 10;
            sheet.Cells.Columns[9].Width = 10;

            obj.FormatCell(sheet.Cells["A2"], "班級");
            obj.FormatCell(sheet.Cells["B2"], "座號");
            obj.FormatCell(sheet.Cells["C2"], "姓名");
            obj.FormatCell(sheet.Cells["D2"], "學號");
            obj.FormatCell(sheet.Cells["E2"], "大功");
            obj.FormatCell(sheet.Cells["F2"], "小功");
            obj.FormatCell(sheet.Cells["G2"], "嘉獎");
            obj.FormatCell(sheet.Cells["H2"], "大過");
            obj.FormatCell(sheet.Cells["I2"], "小過");
            obj.FormatCell(sheet.Cells["J2"], "警告");
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
                //將統計相換算成比值的基底
                int MeritA = 0;
                int MeritB = 0;
                int MeritC = 0;
                int DemeritA = 0;
                int DemeritB = 0;
                int DemeritC = 0;

                if (_sc._selectMode == SelectMode.依日期)
                {
                    //獎勵資料依日期
                    if (_tb.DicByMerit.ContainsKey(student.ID))
                    {
                        foreach (MeritRecord each in _tb.DicByMerit[student.ID])
                        {
                            int A = each.MeritA.HasValue ? each.MeritA.Value : 0;
                            int B = each.MeritB.HasValue ? each.MeritB.Value : 0;
                            int C = each.MeritC.HasValue ? each.MeritC.Value : 0;
                            MeritA += A;
                            MeritB += B;
                            MeritC += C;
                        }
                    }

                    if (_tb.DicByDemerit.ContainsKey(student.ID))
                    {
                        foreach (DemeritRecord each in _tb.DicByDemerit[student.ID])
                        {
                            int A = each.DemeritA.HasValue ? each.DemeritA.Value : 0;
                            int B = each.DemeritB.HasValue ? each.DemeritB.Value : 0;
                            int C = each.DemeritC.HasValue ? each.DemeritC.Value : 0;
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
                        foreach (AutoSummaryRecord summary in _tb.DicByInitialSummary[student.ID])
                        {
                            int mA = summary.MeritA;
                            int mB = summary.MeritB;
                            int mC = summary.MeritC;
                            MeritA += mA;
                            MeritB += mB;
                            MeritC += mC;

                            int demA = summary.DemeritA;
                            int demB = summary.DemeritB;
                            int demC = summary.DemeritC;
                            DemeritA += demA;
                            DemeritB += demB;
                            DemeritC += demC;
                        }
                    }
                }

                _tb.studentUbeIDList.Add(student.ID);

                int rowIndex = index + 2;

                obj.FormatCell(sheet.Cells["A" + rowIndex], student.Class.Name);
                obj.FormatCell(sheet.Cells["B" + rowIndex], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                obj.FormatCell(sheet.Cells["C" + rowIndex], student.Name);
                obj.FormatCell(sheet.Cells["D" + rowIndex], student.StudentNumber);
                obj.FormatCell(sheet.Cells["E" + rowIndex], MeritA.ToString());
                obj.FormatCell(sheet.Cells["F" + rowIndex], MeritB.ToString());
                obj.FormatCell(sheet.Cells["G" + rowIndex], MeritC.ToString());
                obj.FormatCell(sheet.Cells["H" + rowIndex], DemeritA.ToString());
                obj.FormatCell(sheet.Cells["I" + rowIndex], DemeritB.ToString());
                obj.FormatCell(sheet.Cells["J" + rowIndex], DemeritC.ToString());
                index++;
            }
            #endregion

            int sheetIndex = book.Worksheets.Add(); //再加一個Sheet
            Worksheet sheet2 = book.Worksheets[sheetIndex];
            sheet2.Name = "累計明細";

            Aspose.Cells.Row row_sheet2 = sheet2.Cells.Rows[0];
            row_sheet2.Height = 30;
            obj.FormatCell_2(sheet2.Cells["A1"], School.ChineseName + "\n累計明細");

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
            obj.FormatCell(sheet2.Cells["K2"], "大過");
            obj.FormatCell(sheet2.Cells["L2"], "小過");
            obj.FormatCell(sheet2.Cells["M2"], "警告");
            obj.FormatCell(sheet2.Cells["N2"], "事由");
            obj.FormatCell(sheet2.Cells["O2"], "登錄日期");
            #endregion

            int ri = 3;

            foreach (StudentRecord student in StudentList)
            {
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

                            //KLM

                            obj.FormatCell(sheet2.Cells["N" + ri], merit.Reason);
                            obj.FormatCell(sheet2.Cells["O" + ri], merit.RegisterDate.HasValue ? merit.RegisterDate.Value.ToShortDateString() : "");

                            ri++;
                        }

                        foreach (DemeritRecord demerit in _tb.DicByDemerit[student.ID])
                        {
                            obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                            obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                            obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                            obj.FormatCell(sheet2.Cells["E" + ri], demerit.SchoolYear.ToString());
                            obj.FormatCell(sheet2.Cells["F" + ri], demerit.Semester.ToString());
                            obj.FormatCell(sheet2.Cells["G" + ri], demerit.OccurDate.ToShortDateString());

                            obj.FormatCell(sheet2.Cells["K" + ri], demerit.DemeritA.HasValue ? demerit.DemeritA.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["L" + ri], demerit.DemeritB.HasValue ? demerit.DemeritB.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["M" + ri], demerit.DemeritC.HasValue ? demerit.DemeritC.Value.ToString() : "");

                            obj.FormatCell(sheet2.Cells["N" + ri], demerit.Reason);
                            obj.FormatCell(sheet2.Cells["O" + ri], demerit.RegisterDate.HasValue ? demerit.RegisterDate.Value.ToShortDateString() : "");

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

                            obj.FormatCell(sheet2.Cells["N" + ri], merit.Reason);
                            obj.FormatCell(sheet2.Cells["O" + ri], merit.RegisterDate.HasValue ? merit.RegisterDate.Value.ToShortDateString() : "");

                            ri++;
                        }
                    }

                    if (_tb.DicByDemerit.ContainsKey(student.ID))
                    {
                        foreach (DemeritRecord demerit in _tb.DicByDemerit[student.ID])
                        {
                            obj.FormatCell(sheet2.Cells["A" + ri], student.Class.Name);
                            obj.FormatCell(sheet2.Cells["B" + ri], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["C" + ri], student.Name);
                            obj.FormatCell(sheet2.Cells["D" + ri], student.StudentNumber);
                            obj.FormatCell(sheet2.Cells["E" + ri], demerit.SchoolYear.ToString());
                            obj.FormatCell(sheet2.Cells["F" + ri], demerit.Semester.ToString());
                            obj.FormatCell(sheet2.Cells["G" + ri], demerit.OccurDate.ToShortDateString());

                            obj.FormatCell(sheet2.Cells["K" + ri], demerit.DemeritA.HasValue ? demerit.DemeritA.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["L" + ri], demerit.DemeritB.HasValue ? demerit.DemeritB.Value.ToString() : "");
                            obj.FormatCell(sheet2.Cells["M" + ri], demerit.DemeritC.HasValue ? demerit.DemeritC.Value.ToString() : "");

                            obj.FormatCell(sheet2.Cells["O" + ri], demerit.Reason);
                            obj.FormatCell(sheet2.Cells["N" + ri], demerit.RegisterDate.HasValue ? demerit.RegisterDate.Value.ToShortDateString() : "");

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
                                if (summary.InitialMeritA + summary.InitialMeritB + summary.InitialMeritC > 0 || summary.InitialDemeritA + summary.InitialDemeritB + summary.InitialDemeritC > 0)
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
                                    obj.FormatCell(sheet2.Cells["K" + ri], "" + summary.InitialDemeritA);
                                    obj.FormatCell(sheet2.Cells["L" + ri], "" + summary.InitialDemeritB);
                                    obj.FormatCell(sheet2.Cells["M" + ri], "" + summary.InitialDemeritC);
                                    obj.FormatCell(sheet2.Cells["N" + ri], "(非明細資料)");
                                    obj.FormatCell(sheet2.Cells["0" + ri], "");

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
                    obj.PrintNow(book, "獎懲總計");
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印獎懲總計,已完成!");
                }
                else
                {
                    MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印獎懲總計,發生錯誤!");
                }
            }
            else
            {
                MsgBox.Show("列印作業已中止!");
            }
        }
    }
}
