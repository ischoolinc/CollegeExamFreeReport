using FISCA;
using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport108
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            //2021-12-07 Cynthia 因108課綱直接切出一版使用
            // 2022-04-14 Cynthia依照高雄小組要求，將報表改為「五專免試入學相關報表」

            #region 聯合
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "資料統計"];
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「聯合」免試入學 - 108課綱適用"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「聯合」免試入學 - 108課綱適用"].Click += delegate
            {
                Report report = new Report();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.權限_聯合免試入學108)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「聯合」免試入學 - 108課綱適用"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「聯合」免試入學 - 108課綱適用"].Enable = false;
                }
            };
            #endregion

            #region 優先
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「優先」免試入學 - 108課綱適用"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「優先」免試入學 - 108課綱適用"].Click += delegate
            {
                Report_priority report = new Report_priority();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.權限_優先免試入學108)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「優先」免試入學 - 108課綱適用"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「優先」免試入學 - 108課綱適用"].Enable = false;
                }
            };
            #endregion

            #region 完全
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「完全」免試入學 - 108課綱適用"].Enable = false;
            item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「完全」免試入學 - 108課綱適用"].Click += delegate
            {
                Report_Completely report = new Report_Completely();
                report.ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.權限_完全免試入學108)
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「完全」免試入學 - 108課綱適用"].Enable = true;
                }
                else
                {
                    item1["報表"]["成績相關報表"]["五專免試入學相關報表"]["超額比序項目積分證明單「完全」免試入學 - 108課綱適用"].Enable = false;
                }
            };
            #endregion

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.聯合免試入學108, "超額比序項目積分證明單「聯合」免試入學 - 108課綱適用"));

            permission.Add(new RibbonFeature(Permissions.優先免試入學108, "超額比序項目積分證明單「優先」免試入學 - 108課綱適用"));

            permission.Add(new RibbonFeature(Permissions.完全免試入學108, "超額比序項目積分證明單「完全」免試入學 - 108課綱適用"));
        }

    }
}
