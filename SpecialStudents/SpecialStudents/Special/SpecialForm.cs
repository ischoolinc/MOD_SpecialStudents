using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.DSAUtil;
using System.Xml;
using Aspose.Cells;
using System.IO;
using System.Diagnostics;
using K12.Data;
using FISCA;

namespace SpecialStudents
{
    public enum SelectMode
    {
        依學期, 所有學期, 依日期
    }

    public partial class SpecialForm : BaseForm
    {
        PrintObj obj = new PrintObj();

        BackgroundWorker BGW = new BackgroundWorker();

        List<string> AttendanceStringList = new List<string>();

        Dictionary<string, string> PeriodTypeDic = new Dictionary<string, string>();

        Dictionary<string, bool> AttendanceIsNoabsence = new Dictionary<string, bool>();

        List<K12.Data.AbsenceMappingInfo> InfoList { get; set; }

        List<K12.Data.PeriodMappingInfo> PerList { get; set; }

        //學生清單
        List<StudentRecord> _StudentRecordList = new List<StudentRecord>();

        K12.Data.Configuration.ConfigData Attcd { get; set; }

        public SpecialForm()
        {
            InitializeComponent();
        }

        //載入預設畫面
        private void SpecialForm_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            tabControl1.Enabled = false;
            this.Text = "學生資料讀取中";

            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            _StudentRecordList = obj.GetStudentList();

            AttendanceIsNoabsence = tool.GetAbsenceMapping();
            PeriodTypeDic = tool.GetPeriodMapping();

            PerList = K12.Data.PeriodMapping.SelectAll();
            Attcd = K12.Data.School.Configuration[tool.AttConfigName];

        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Text = "查詢學生特殊表現名單";
            tabControl1.Enabled = true;
            SetSchoolYearSemester();

            SetForm1();
        }

        //預設畫面1內容(累計節次,假別)
        private void SetForm1()
        {
            txtPeriodCount.Text = "1";

            //缺曠別
            AttendanceStringList.Clear();
            listViewEx1.Items.Clear();

            dateTimeInput1.Value = DateTime.Today.AddDays(-7);
            dateTimeInput2.Value = DateTime.Today;

            foreach (string e in AttendanceIsNoabsence.Keys)
            {
                AttendanceStringList.Add(e);
                listViewEx1.Items.Add(e);
            }
        }

        //列印"缺曠累計名單"
        private void btnPrint1_Click(object sender, EventArgs e)
        {
            AttendanceScClick Atsc = new AttendanceScClick();

            SetConfig _sc = DefSetup();

            Atsc.print(_sc, txtPeriodCount, listViewEx1, AttendanceStringList, PeriodTypeDic);

        }

        //列印"全勤學生"名單
        private void btnPrint2_Click(object sender, EventArgs e)
        {
            NoAbsenceScClick NAsc = new NoAbsenceScClick();
            NAsc.AttendanceIsNoabsence = AttendanceIsNoabsence;

            SetConfig _sc = DefSetup();

            NAsc.print(_sc);
        }

        //列印"獎勵特殊表現"名單
        private void btnPrint4_Click(object sender, EventArgs e)
        {
            MeritScClick Msc = new MeritScClick();

            SetConfig _sc = DefSetup();

            Msc.print(_sc, tbMeritA, tbMeritB, tbMeritC, cbxIgnoreDemerit, cbxDemeritIsNull, cbxIsDemeritClear);
        }

        //列印"懲戒特殊表現"名單
        private void btnPrint3_Click(object sender, EventArgs e)
        {
            DemeritScClick Dmsc = new DemeritScClick();

            SetConfig _sc = DefSetup();

            Dmsc.Print(_sc, tbDemeritA, tbDemeritB, tbDemeritC, cbxIsMeritAndDemerit);
        }

        private SetConfig DefSetup()
        {
            SelectMode _Select = SelectMode.依學期;
            if (checkBoxX3.Checked)
                _Select = SelectMode.所有學期;
            else if (checkBoxX4.Checked)
                _Select = SelectMode.依日期;

            SetConfig _sc = new SetConfig();
            _sc._SchoolYear = intSchoolYear1.Value;
            _sc._Semester = intSemester1.Value;
            _sc._StudentList = _StudentRecordList;
            _sc._selectMode = _Select;
            _sc.StartDate = dateTimeInput1.Value;
            _sc.EndDate = dateTimeInput2.Value;

            return _sc;
        }

        //預設畫面的學年度學期
        private void SetSchoolYearSemester()
        {
            int SchoolYear;
            int Semester;

            if (int.TryParse(School.DefaultSchoolYear, out SchoolYear))
            {
                intSchoolYear1.Value = SchoolYear;
            }

            if (int.TryParse(School.DefaultSemester, out Semester))
            {
                intSemester1.Value = Semester;
            }

        }

        //全選假別內容
        private void cbxSelectAllPeriod_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem each in listViewEx1.Items)
            {
                each.Checked = cbxSelectAllPeriod.Checked;
            }
        }

        //離開本功能
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Link


        string URL缺曠類別管理 = "ischool/國中系統/學務/管理/缺曠類別管理";

        string URL功過換算管理 = "ischool/國中系統/學務/管理/功過換算管理";

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Features.Invoke(URL缺曠類別管理);
            }
            catch
            {
                MsgBox.Show("無此功能!");
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Features.Invoke(URL功過換算管理);
            }
            catch
            {
                MsgBox.Show("無此功能!");
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Features.Invoke(URL功過換算管理);
            }
            catch
            {
                MsgBox.Show("無此功能!");
            }
        }
        #endregion

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BalanceConfigForm BCForm = new BalanceConfigForm();
            BCForm.ShowDialog();
        }
    }
}
