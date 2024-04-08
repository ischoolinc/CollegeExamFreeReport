using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport108
{
    class Permissions
    {
        public static string 聯合免試入學108 { get { return "CollegeExamFreeReport108.46CE921D-338A-4980-86BB-2E6C39FBCFA5"; } }

        public static string 優先免試入學108 { get { return "CollegeExamFreeReport108.B91BA7CC-22BE-4B5D-AA34-E5142257BD8A"; } }

        public static string 完全免試入學108 { get { return "CollegeExamFreeReport108.0346474f-135a-4309-8fe4-6522a0fc1f86"; } }

        public static bool 權限_聯合免試入學108
        {
            get { return FISCA.Permission.UserAcl.Current[聯合免試入學108].Executable; }
        }

        public static bool 權限_優先免試入學108
        {
            get { return FISCA.Permission.UserAcl.Current[優先免試入學108].Executable; }
        }

        public static bool 權限_完全免試入學108
        {
            get { return FISCA.Permission.UserAcl.Current[完全免試入學108].Executable; }
        }
    }
}
