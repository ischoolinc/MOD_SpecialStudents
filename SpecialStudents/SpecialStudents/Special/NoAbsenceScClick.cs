using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using K12.Data;
using DevComponents.DotNetBar.Controls;
using DevComponents.Editors;
using System.Drawing;
using System.ComponentModel;
using FISCA.Presentation.Controls;
using K12.BusinessLogic;

namespace SpecialStudents
{
    class NoAbsenceScClick
    {
        PrintObj obj;

        BackgroundWorker BGW = new BackgroundWorker();

        //報表專用
        Workbook book;

        StudentAbsence _sa { get; set; }

        //系統缺曠別
        public Dictionary<string, bool> _AttendanceIsNoabsence = new Dictionary<string, bool>();

        SetConfig _sc { get; set; }

        public void print(SetConfig sc, Dictionary<string, bool> AttendanceIsNoabsence)
        {
            _sc = sc;
            _AttendanceIsNoabsence = AttendanceIsNoabsence;

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("全勤學生清單,列印中!");
            BGW.RunWorkerAsync();

        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();

            List<string> _StudentIDList = _sc.GetStudentIdList();

            if (_sc._selectMode == SelectMode.所有學期)
            {
                //取得AutoSummary
                List<AutoSummaryRecord> AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Attendance);

                _sa = GetSABySchoolYear(AutoSummaryList);


            }
            else if (_sc._selectMode == SelectMode.依學期)
            {
                //取得AutoSummary
                SchoolYearSemester SYS = new SchoolYearSemester(_sc._SchoolYear, _sc._Semester);
                List<AutoSummaryRecord> AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Attendance);

                _sa = GetSABySchoolYear(AutoSummaryList);
            }
            else
            {
                _sa = GetSAByDate();
            }

            //排除有影響的學生就是我要的學生
            foreach (StudentRecord each in _sc._StudentList) //
            {
                if (!_sa.StudentAbsenceList.Contains(each.ID)) //不包含在有資料清單內
                {
                    _sa.StudentNoAbsenceList.Add(each.ID); //有全勤的學生
                }
            }

            book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.Name = "全勤學生清單";

            Cell A1 = sheet.Cells["A1"];
            A1 = tool.UserStyle(A1);

            string A1Name = School.ChineseName + "　全勤學生清單";
            if (_sc._selectMode == SelectMode.依學期)
            {
                A1Name += "　(" + _sc._SchoolYear.ToString() + " / " + _sc._Semester.ToString() + ")";
            }
            else if (_sc._selectMode == SelectMode.依日期)
            {
                A1Name += "　(" + _sc.StartDate.ToShortDateString() + " ~ " + _sc.EndDate.ToShortDateString() + ")";
            }
            else
            {
                A1Name += "(所有學期)";
            }

            //A1.PutValue(A1Name);
            obj.FormatCell(A1, A1Name);
            //sheet.Cells.Merge(0, 0, 1, 5);

            obj.FormatCell(sheet.Cells["A2"], "編號");
            obj.FormatCell(sheet.Cells["B2"], "班級");
            obj.FormatCell(sheet.Cells["C2"], "座號");
            obj.FormatCell(sheet.Cells["D2"], "姓名");
            obj.FormatCell(sheet.Cells["E2"], "學號");

            int index = 1;

            List<StudentRecord> studentList = Student.SelectByIDs(_sa.StudentNoAbsenceList);

            studentList = SortClassIndex.K12Data_StudentRecord(studentList);

            foreach (StudentRecord each in studentList)
            {
                int rowIndex = index + 2;
                obj.FormatCell(sheet.Cells["A" + rowIndex], index.ToString());
                obj.FormatCell(sheet.Cells["B" + rowIndex], each.Class.Name);
                obj.FormatCell(sheet.Cells["C" + rowIndex], each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "");
                obj.FormatCell(sheet.Cells["D" + rowIndex], each.Name);
                obj.FormatCell(sheet.Cells["E" + rowIndex], each.StudentNumber);
                index++;
            }
        }

        private StudentAbsence GetSAByDate()
        {
            StudentAbsence sa = new StudentAbsence();
            List<AttendanceRecord> AttendList = K12.Data.Attendance.SelectByDate(_sc._StudentList, _sc.StartDate, _sc.EndDate);

            foreach (AttendanceRecord ar in AttendList)
            {
                if (ar.PeriodDetail.Count == 0)
                    continue;

                foreach (AttendancePeriod perid in ar.PeriodDetail)
                {
                    //是否影響全勤判斷
                    //不包含假別中就離開
                    if (!_AttendanceIsNoabsence.ContainsKey(perid.AbsenceType))
                        continue;

                    if (!_AttendanceIsNoabsence[perid.AbsenceType]) //True就是不影響全勤
                    {
                        if (!sa.StudentAbsenceList.Contains(ar.RefStudentID)) //如果沒有就加入
                        {
                            sa.StudentAbsenceList.Add(ar.RefStudentID);
                            break;
                        }
                    }
                }
            }


            return sa;
        }

        /// <summary>
        /// 取得缺曠學生
        /// 所有學期 or 單一學期
        /// </summary>
        /// <param name="AutoSummaryList"></param>
        /// <returns></returns>
        private StudentAbsence GetSABySchoolYear(List<AutoSummaryRecord> AutoSummaryList)
        {
            StudentAbsence sa = new StudentAbsence();

            foreach (AutoSummaryRecord each in AutoSummaryList)
            {
                foreach (AbsenceCountRecord count in each.AbsenceCounts)
                {
                    if (count.Count == 0)
                        continue;

                    if (!_AttendanceIsNoabsence.ContainsKey(count.Name)) //不包含假別中就離開
                        continue;

                    if (!_AttendanceIsNoabsence[count.Name]) //True就是不影響全勤
                    {
                        if (!sa.StudentAbsenceList.Contains(each.RefStudentID)) //如果沒有就加入
                        {
                            sa.StudentAbsenceList.Add(each.RefStudentID);
                        }
                    }
                }
            }

            return sa;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    obj.PrintNow(book, "全勤學生清單");
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印全勤學生清單,已完成!");
                }
                else
                {
                    MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印全勤學生清單,發生錯誤!");
                }
            }
            else
            {
                MsgBox.Show("列印作業已中止!");
            }
        }
    }
}
