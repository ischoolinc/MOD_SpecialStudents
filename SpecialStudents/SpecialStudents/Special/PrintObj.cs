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
            foreach (Worksheet sheet in _book.Worksheets)
            {
                sheet.AutoFitColumns();
            }

            string path = Path.Combine(Application.StartupPath, "Reports");

            //如果目錄不存在則建立。
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            path = Path.Combine(path, Name + ".xls");
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
                    string nameTempalte = file.FullName.Replace(file.Extension, "") + "{0}.xls";
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
        /// 傳入3個元件檢查是否輸入為數字(true為錯誤),
        /// </summary>
        public bool CheckTextBox(string a1, string b1, string c1)
        {
            int TextString;

            //不是數字 and 不是空字串
            if (!int.TryParse(a1, out TextString))
            {
                return true;
            }

            if (!int.TryParse(b1, out TextString))
            {
                return true;
            }

            if (!int.TryParse(c1, out TextString))
            {
                return true;
            }

            return false;

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
            List<K12.Data.StudentRecord> StudentRecordList = new List<K12.Data.StudentRecord>();

            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;

            foreach (K12.Data.StudentRecord student in K12.Data.Student.SelectAll())
            {
                //判斷學生狀態
                if (student.Status != K12.Data.StudentRecord.StudentStatus.一般)
                    continue;

                if (student.Class == null)
                    continue;

                //是選取班級之學生
                string classid = student.Class != null ? student.Class.ID : "";
                if (ClassIDList.Contains(classid))
                {
                    if (!StudentRecordList.Contains(student))
                    {
                        StudentRecordList.Add(student);
                    }
                }
            }

            //排序
            //StudentRecordList.Sort(new SortClass().SortStudent);
            //StudentRecordList = SortClassIndex.JHSchoolData_JHStudentRecord(StudentRecordList);
            return StudentRecordList;
        }
    }
}
