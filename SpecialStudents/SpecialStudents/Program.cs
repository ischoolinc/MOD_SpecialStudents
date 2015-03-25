using FISCA;
using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecialStudents
{
    public class Program
    {
        [MainMethod()]
        static public void Main()
        {
            RibbonBarItem classSpecialItem = K12.Presentation.NLDPanels.Class.RibbonBarItems["學務"];

            classSpecialItem["特殊表現學生"].Size = RibbonBarButton.MenuButtonSize.Large;
            classSpecialItem["特殊表現學生"].Image = Properties.Resources.attendance_list_ok_64;
            classSpecialItem["特殊表現學生"].Enable = false;
            classSpecialItem["特殊表現學生"].Click += delegate
            {
                SpecialForm SpecialFormNew = new SpecialForm();
                SpecialFormNew.ShowDialog();
            };

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                classSpecialItem["特殊表現學生"].Enable = (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 && Permissions.特殊表現學生權限);
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("JHSchool.Class.Ribbon0060", "特殊表現學生"));

        }
    }
}
