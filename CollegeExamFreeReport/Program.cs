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
        {   // [ischoolkingdom] Vicky更動，原107年五專(優先)免試入學相關報表，移除"107年"字樣
            //2021-01-25 Cynthia依照要求將一般免試入學修改成聯合免試入學
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(聯合免試入學)"].Enable = false;
            item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(聯合免試入學)"].Click += delegate
            {
                Report report = new Report();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.超額比序項目積分證明單權限_一般免試入學)
                {
                    item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(聯合免試入學)"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(聯合免試入學)"].Enable = false;
                }
            };


            item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(優先免試入學)"].Enable = false;
            item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(優先免試入學)"].Click += delegate
            {
                Report_priority report = new Report_priority();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.超額比序項目積分證明單權限_優先免試入學)
                {
                    item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(優先免試入學)"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專(優先)免試入學相關報表"]["超額比序項目積分證明單(優先免試入學)"].Enable = false;
                }
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.超額比序項目積分證明單_一般免試入學, "超額比序項目積分證明單(聯合免試入學)"));

            permission.Add(new RibbonFeature(Permissions.超額比序項目積分證明單_優先免試入學, "超額比序項目積分證明單(優先免試入學)"));
        }

    }
}
