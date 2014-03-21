using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport
{
    class Permissions
    {
        public static string 超額比序項目積分證明單 { get { return "CollegeExamFreeReport.46CE921D-338A-4980-86BB-2E6C39FBCFA5"; } }

        public static bool 超額比序項目積分證明單權限
        {
            get { return FISCA.Permission.UserAcl.Current[超額比序項目積分證明單].Executable; }
        }
    }
}
