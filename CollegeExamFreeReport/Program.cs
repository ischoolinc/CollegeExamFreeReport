using FISCA;
using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單"].Click += delegate
            {
                Report report = new Report();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.超額比序項目積分證明單權限)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單"].Enable = false;
                }
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.超額比序項目積分證明單, "超額比序項目積分證明單"));
        }

    }
}
