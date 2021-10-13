using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Cells;
using System.Windows.Forms;
using System.Diagnostics;
using DevComponents.DotNetBar.Controls;
using System.Drawing;
using System.Data;

namespace SpecialStudents
{
    class PrintObj
    {
        //報表專用
        private Workbook _book;

        /// <summary>
        /// 產生報表/列印報表
        /// </summary>
        public void PrintNow(Workbook book, string Name)
        {
            _book = book;

            //foreach (Worksheet sheet in _book.Worksheets)
            //{
            //    sheet.AutoFitRows();
            //    sheet.AutoFitColumns();
            //}

            string path = Path.Combine(Application.StartupPath, "Reports");

            //如果目錄不存在則建立。
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            path = Path.Combine(path, Name + ".xlsx");
            int i = 1;
            while (true)
            {
                string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                if (!File.Exists(newPath))
                {
                    path = newPath;
                    break;
                }
            }
            try
            {
                _book.Save(path);
            }
            catch (IOException)
            {
                try
                {
                    FileInfo file = new FileInfo(path);
                    string nameTempalte = file.FullName.Replace(file.Extension, "") + "{0}.xlsx";
                    int count = 1;
                    string fileName = string.Format(nameTempalte, count);
                    while (File.Exists(fileName))
                        fileName = string.Format(nameTempalte, count++);

                    _book.Save(fileName);
                    path = fileName;
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(path);
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案開啟失敗:" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// 清空特殊符號
        /// </summary>
        private string ConvertToValidName(string A1Name)
        {
            char[] invalids = Path.GetInvalidFileNameChars();

            string result = A1Name;
            foreach (char each in invalids)
                result = result.Replace(each, '_');

            return result;
        }

        /// <summary>
        /// 如果是double就乘上基數
        /// </summary>
        public bool doubleCheck(string txt)
        {
            double NowValue;
            if (!double.TryParse(txt, out NowValue))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 格式化Cell
        /// </summary>
        public void FormatCell(Cell cell, string value)
        {
            cell.PutValue(value);

            Style style = cell.GetStyle();

            style.Borders.SetStyle(CellBorderType.Hair);
            style.Borders.SetColor(Color.Black);
            style.Borders.DiagonalStyle = CellBorderType.None;
            style.HorizontalAlignment = TextAlignmentType.Center;

            cell.SetStyle(style);
        }

        //取得選取班級之學生
        public List<K12.Data.StudentRecord> GetStudentList()
        {
            List<string> StudentIDList = new List<string>();

            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;

            DataTable dt = tool._Q.Select(string.Format("select student.id,student.ref_class_id from student where ref_class_id is not null and ref_class_id in ('{0}') and status in ('1','2')", string.Join("','", ClassIDList)));

            foreach (DataRow row in dt.Rows)
            {
                //是選取班級之學生

                if (!StudentIDList.Contains("" + row["id"]))
                {
                    StudentIDList.Add("" + row["id"]);
                }
            }

            //排序
            //StudentRecordList.Sort(new SortClass().SortStudent);
            //StudentRecordList = SortClassIndex.JHSchoolData_JHStudentRecord(StudentRecordList);
            return K12.Data.Student.SelectByIDs(StudentIDList);
        }
    }
}
