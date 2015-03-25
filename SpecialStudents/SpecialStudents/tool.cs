using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    static public class tool
    {
        /// <summary>
        /// 缺曠類別權重設定 - 功能代碼
        /// </summary>
        static public string AttConfigName = "特殊學生表現_缺曠累積名單";

        static public FISCA.Data.QueryHelper _Q = new FISCA.Data.QueryHelper();

        static public Cell UserStyle(Cell cell)
        {
            Style style = cell.GetStyle();
            style.Borders.SetColor(Color.Black);
            style.HorizontalAlignment = TextAlignmentType.Center;
            cell.SetStyle(style);

            return cell;

        }

        static public Dictionary<string, bool> GetAbsenceMapping()
        {
            Dictionary<string, bool> dic = new Dictionary<string, bool>();

            List<K12.Data.AbsenceMappingInfo> InfoList = K12.Data.AbsenceMapping.SelectAll();

            foreach (K12.Data.AbsenceMappingInfo e in InfoList)
            {
                if (!dic.ContainsKey(e.Name))
                {
                    dic.Add(e.Name, e.Noabsence);
                }
            }
            return dic;
        }

        static public Dictionary<string, string> GetPeriodMapping()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            List<K12.Data.PeriodMappingInfo> InfoList = K12.Data.PeriodMapping.SelectAll();

            foreach (K12.Data.PeriodMappingInfo e in InfoList)
            {
                if (!dic.ContainsKey(e.Name))
                {
                    dic.Add(e.Name, e.Type);
                }
            }

            return dic;
        }
    }
}
