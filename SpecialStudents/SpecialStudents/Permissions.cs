using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    class Permissions
    {
        public static string 特殊表現學生 { get { return "JHSchool.Class.Ribbon0060"; } }
        public static bool 特殊表現學生權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[特殊表現學生].Executable;
            }
        }
    }
}
