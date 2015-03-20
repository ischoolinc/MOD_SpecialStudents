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

namespace SpecialStudents
{
    public partial class BalanceConfigForm : BaseForm
    {
        K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[tool.AttConfigName];

        public BalanceConfigForm()
        {
            InitializeComponent();
        }

        private void BalanceConfigForm_Load(object sender, EventArgs e)
        {
            //取得每日節次對照表(設定畫面)
            List<string> list = new List<string>();

            Dictionary<string, string> PeriodTypeDic = tool.GetPeriodMapping();

            foreach (string info in PeriodTypeDic.Values)
            {
                if (!list.Contains(info))
                {
                    list.Add(info);
                }
            }

            list.Sort();

            //取得設定檔
            Dictionary<string, double> ConfigByName = GetConfigData(list);

            foreach (string each in ConfigByName.Keys)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Cells[0].Value = each;
                row.Cells[1].Value = ConfigByName[each]; //如果沒有設定值...預設為1
                dataGridViewX1.Rows.Add(row);
            }

        }

        private Dictionary<string, double> GetConfigData(List<string> list)
        {
            Dictionary<string, double> ConfigByName = new Dictionary<string, double>();

            if (cd.Count != 0) 
            {
                foreach (string each in list) //對缺曠假別做連集,預設為1
                {
                    if (cd.Contains(each))
                    {
                        if (doubleCheck(cd[each]))
                        {
                            ConfigByName.Add(each, double.Parse(cd[each]));
                        }
                        else
                        {
                            MsgBox.Show(each + "之權重設定目前[" + cd[each] + "],是錯誤狀態,已預設為1");

                            ConfigByName.Add(each, 1);
                        }
                    }
                    else
                    {
                        ConfigByName.Add(each, 1);
                    }
                }
            }
            else //如果沒有設定
            {
                foreach (string each in list)
                {
                    ConfigByName.Add(each, 1);
                }
            }

            return ConfigByName;

        }

        //儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {
            //先移除
            //K12.Data.School.Configuration.Remove(cd);

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {

                if (row.Cells[1].ErrorText != "")
                {
                    MsgBox.Show("輸入資料有誤,請修正後再儲存!");
                    return;
                }

                string Cell1 = ""+row.Cells[0].Value;
                string Cell2 = ""+row.Cells[1].Value;

                cd[Cell1] = Cell2;

            }

            cd.Save();

            FISCA.Presentation.Controls.MsgBox.Show("儲存設定成功");
            this.Close();
        }

        //離開
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridViewX1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string Cellvalue = "" + dataGridViewX1.CurrentCell.Value;

            if (doubleCheck(Cellvalue))
            {
                dataGridViewX1.CurrentCell.ErrorText = "";
            }
            else
            {
                dataGridViewX1.CurrentCell.ErrorText = "儲存格必須輸入數字";
            }
        }

        private bool doubleCheck(string txt)
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
    }
}
