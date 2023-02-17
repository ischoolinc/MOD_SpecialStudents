using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using DevComponents.DotNetBar;
using DevComponents.Editors;
using DevComponents.DotNetBar.Controls;
using FISCA.Presentation.Controls;
using System.Windows.Forms;
using Aspose.Cells;
using System.ComponentModel;
using K12.BusinessLogic;
using System.Drawing;

namespace SpecialStudents
{
    class AttendanceScClick
    {
        //自動統計內容
        List<AutoSummaryRecord> AutoSummaryList;

        BackgroundWorker BGW = new BackgroundWorker();
        //
        PrintObj obj { get; set; }

        //系統的缺曠別清單
        public List<string> _AttendaceName { get; set; }

        //報表專用
        Workbook book;

        //累計節次
        double PeriodCount;

        K12.Data.Configuration.ConfigData cd { get; set; }

        //缺曠清單
        List<string> SelectAbsenceList = new List<string>();

        Dictionary<string, string> _PeriodTypeDic { get; set; }

        //缺曠統計&權重比
        Dictionary<string, Dictionary<string, double>> studentAttendance2 { get; set; }

        SetConfig _sc { get; set; }

        public void print(SetConfig sc, TextBoxX txtPeriodCount, ListViewEx listViewEx1, List<string> AttendaceName, Dictionary<string, string> PeriodTypeDic)
        {
            _sc = sc;
            _AttendaceName = AttendaceName;
            _PeriodTypeDic = PeriodTypeDic;

            if (!double.TryParse(txtPeriodCount.Text, out PeriodCount))
            {
                MsgBox.Show("累計節次內容非數字!!");
                SpecialEvent.RaiseSpecialChanged();
                return;
            }

            //選擇的缺曠類別
            SelectAbsenceList.Clear();
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    SelectAbsenceList.Add(item.Text);
                }
            }

            if (SelectAbsenceList.Count == 0)
            {
                MsgBox.Show("至少必須選擇一個缺曠類別!");
                SpecialEvent.RaiseSpecialChanged();
                return;
            }

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            FISCA.Presentation.MotherForm.SetStatusBarMessage("缺曠累計名單,列印中!");
            BGW.RunWorkerAsync();
        }
        //開始背景模式
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            obj = new PrintObj();
            cd = K12.Data.School.Configuration[tool.AttConfigName];

            //取得缺曠內容
            //List<JHAttendanceRecord> AttendanceList;

            List<string> _StudentIDList = _sc.GetStudentIdList();

            if (_sc._selectMode == SelectMode.所有學期)
            {
                List<AutoSummaryRecord> AutoSummaryList = AutoSummary.Select(_StudentIDList, new List<SchoolYearSemester>(), SummaryType.Attendance);
                studentAttendance2 = GetAttendanceDetail(AutoSummaryList);
            }
            else if (_sc._selectMode == SelectMode.依學期)
            {
                SchoolYearSemester SYS = new SchoolYearSemester(_sc._SchoolYear, _sc._Semester);
                List<AutoSummaryRecord> AutoSummaryList = AutoSummary.Select(_StudentIDList, new SchoolYearSemester[] { SYS }, SummaryType.Attendance);
                studentAttendance2 = GetAttendanceDetail(AutoSummaryList);
            }
            else //依日期
            {
                studentAttendance2 = GetDayCountDetail();
            }

            #region 預設報表資訊
            book = new Workbook();
            book.Worksheets.Clear();

            int sheetIndex = book.Worksheets.Add();
            Worksheet sheet = book.Worksheets[sheetIndex];
            sheet.Name = "缺曠累計名單";

            //將格子合併
            sheet.Cells.Merge(0, 0, 1, 5 + SelectAbsenceList.Count);

            string A1Name = School.ChineseName + "\n缺曠累計名單";
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


            //sheet.Cells[0, 0].PutValue(A1Name);

            Aspose.Cells.Row row = sheet.Cells.Rows[0];
            row.Height = 30;
            obj.FormatCell_2(sheet.Cells[0, 0], A1Name);

            obj.FormatCell(sheet.Cells[1, 0], "班級");
            obj.FormatCell(sheet.Cells[1, 1], "座號");
            obj.FormatCell(sheet.Cells[1, 2], "姓名");
            obj.FormatCell(sheet.Cells[1, 3], "學號");

            Dictionary<string, int> saveAttAddress1 = new Dictionary<string, int>();

            int countList = 4;
            foreach (string var in SelectAbsenceList) //依選擇的假別
            {
                saveAttAddress1.Add(var, countList); //記錄定位
                obj.FormatCell(sheet.Cells[1, countList], var);

                countList++;
            }
            obj.FormatCell(sheet.Cells[1, countList], "累積節次");

            int cellcount = 2; //ROW的Index
            //int _MergeInt = 0; 
            #endregion

            List<StudentRecord> PrintStudentList = Student.SelectByIDs(studentAttendance2.Keys);
            PrintStudentList = SortClassIndex.K12Data_StudentRecord(PrintStudentList);

            #region 學生逐一列印
            //取得一名學生之資料
            foreach (StudentRecord student in PrintStudentList)
            {
                string var = student.ID;

                double xyz = 0;
                //處理假別相加
                foreach (string invar in studentAttendance2[var].Keys)
                {
                    //假別是否是使用者所選
                    if (SelectAbsenceList.Contains(invar))
                    {
                        //將資料相加
                        xyz = xyz + studentAttendance2[var][invar];
                    }
                }

                //如果累計數量大於等於使用者所輸入
                if (xyz >= PeriodCount)
                {
                    //班級
                    obj.FormatCell(sheet.Cells[cellcount, 0], student.Class.Name);

                    //座號
                    obj.FormatCell(sheet.Cells[cellcount, 1], student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "");

                    //姓名
                    obj.FormatCell(sheet.Cells[cellcount, 2], student.Name);

                    //學號
                    obj.FormatCell(sheet.Cells[cellcount, 3], student.StudentNumber);

                    //權重累積
                    obj.FormatCell(sheet.Cells[cellcount, countList], "" + xyz);

                    //數量
                    foreach (string invar in studentAttendance2[var].Keys) //取得假別
                    {
                        if (saveAttAddress1.ContainsKey(invar)) //是否有在定位內
                        {
                            obj.FormatCell(sheet.Cells[cellcount, saveAttAddress1[invar]], ""+studentAttendance2[var][invar]);
                        }
                    }

                    //補0機制...
                    foreach (int injar in saveAttAddress1.Values)
                    {
                        if (sheet.Cells[cellcount, injar].StringValue == string.Empty)
                        {
                            sheet.Cells[cellcount, injar].PutValue(0);
                        }
                        //_MergeInt = injar;
                    }

                    cellcount++;
                }
            }
            #endregion
        }

        private Dictionary<string, Dictionary<string, double>> GetDayCountDetail()
        {
            //學生ID : 缺曠別 : 比值
            Dictionary<string, Dictionary<string, double>> dic = new Dictionary<string, Dictionary<string, double>>();
            List<AttendanceRecord> AttendList = K12.Data.Attendance.SelectByDate(_sc._StudentList, _sc.StartDate, _sc.EndDate);
            Dictionary<string, double> cdName = new Dictionary<string, double>();
            foreach (string each in cd)
            {
                if (!cdName.ContainsKey(each))
                {
                    cdName.Add(each, double.Parse(cd[each]));
                }
            }


            foreach (AttendanceRecord each in AttendList)
            {
                if (!dic.ContainsKey(each.RefStudentID))
                {
                    dic.Add(each.RefStudentID, new Dictionary<string, double>());
                }

                foreach (AttendancePeriod per in each.PeriodDetail)
                {
                    //如果不是系統中假別及節次,不予計算
                    if (!_AttendaceName.Contains(per.AbsenceType))
                        continue;

                    if (!dic[each.RefStudentID].ContainsKey(per.AbsenceType))
                    {
                        dic[each.RefStudentID].Add(per.AbsenceType, 0); //預設為0
                    }

                    //是否為目前節次清單內
                    if (_PeriodTypeDic.ContainsKey(per.Period))
                    {
                        string perTyp = _PeriodTypeDic[per.Period];

                        //是否為設定中的節次類型 : 一般或集會
                        if (cdName.ContainsKey(perTyp))
                        {
                            dic[each.RefStudentID][per.AbsenceType] += cdName[perTyp];
                        }
                        else
                        {
                            dic[each.RefStudentID][per.AbsenceType] += 1;
                        }
                    }
                }
            }

            return dic;
        }

        /// <summary>
        /// 依據AutoSummary(依單學期&依所有學期)
        /// 取得學生缺曠統計狀況
        /// </summary>
        private Dictionary<string, Dictionary<string, double>> GetAttendanceDetail(List<AutoSummaryRecord> AutoSummaryList)
        {

            Dictionary<string, Dictionary<string, double>> dic = new Dictionary<string, Dictionary<string, double>>();

            foreach (AutoSummaryRecord auto in AutoSummaryList)
            {
                //字典是否有此學生
                if (!dic.ContainsKey(auto.RefStudentID))
                {
                    dic.Add(auto.RefStudentID, new Dictionary<string, double>());
                }

                foreach (AbsenceCountRecord absence in auto.AbsenceCounts)
                {
                    //如果不是系統中假別及節次,不予計算
                    if (!_AttendaceName.Contains(absence.Name))
                        continue;

                    //此學生字典是否已有此假別
                    if (!dic[auto.RefStudentID].ContainsKey(absence.Name))
                    {
                        dic[auto.RefStudentID].Add(absence.Name, 0); //預設為0
                    }

                    //記權重(進行換算)
                    if (cd.Contains(absence.PeriodType))
                    {
                        if (obj.doubleCheck(cd[absence.PeriodType])) //如果是double就乘上基數
                        {
                            dic[auto.RefStudentID][absence.Name] += absence.Count * double.Parse(cd[absence.PeriodType]);
                        }
                        else //不是就預設為1
                        {
                            dic[auto.RefStudentID][absence.Name] += absence.Count * 1;
                        }
                    }
                    else
                    {
                        dic[auto.RefStudentID][absence.Name] += absence.Count * 1;
                    }
                }
            }

            return dic;
        }

        //列印完成
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SpecialEvent.RaiseSpecialChanged();
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    obj.PrintNow(book, "缺曠累計名單");
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印缺曠累計名單,已完成!");
                }
                else
                {
                    MsgBox.Show("列印時發生錯誤!!" + e.Error.Message);
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("列印缺曠累計名單,發生錯誤!");

                }
            }
            else
            {
                MsgBox.Show("列印作業已中止!");
            }
        }
    }
}
