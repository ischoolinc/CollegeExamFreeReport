using Aspose.Words;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using JHSchool.Data;
using K12.BusinessLogic;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CollegeExamFreeReport
{
    public partial class Report : BaseForm
    {
        Configure _Configure;
        AccessHelper _A = new AccessHelper();
        QueryHelper _Q = new QueryHelper();
        BackgroundWorker _BW;
        string _SchoolName,_SchoolCode;
        Dictionary<String, String> _Column2Items;
        Dictionary<String, List<string>> _MappingData;

        public Report()
        {
            InitializeComponent();
            Column1Prepare();
            Column2Prepare();
            _SchoolName = K12.Data.School.ChineseName;
            _SchoolCode = K12.Data.School.Code;
            _BW = new BackgroundWorker();
            _BW.WorkerReportsProgress = true;
            _BW.DoWork += new DoWorkEventHandler(DataBuilding);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ReportBuilding);
            _BW.ProgressChanged += new ProgressChangedEventHandler(BW_Progress);
        }

        private void BW_Progress(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(Global.ReportName + "產生中", e.ProgressPercentage);
        }

        private void Column1Prepare()
        {
            this.Column1.Items.Add("低收入戶");
            this.Column1.Items.Add("中收入戶");
            this.Column1.Items.Add("直系血親尊親屬支領失業給付");
            this.Column1.Items.Add("特殊境遇家庭");
        }

        private void Column2Prepare()
        {
            _Column2Items = new Dictionary<String, String>();

            DataTable dt = _Q.Select("SELECT * FROM tag WHERE category='Student' ORDER BY prefix,name");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();

                string key = "";
                if (string.IsNullOrWhiteSpace(prefix))
                {
                    key = name;
                }
                else
                {
                    key = prefix + ":" + name;
                }

                if (!_Column2Items.ContainsKey(key))
                {
                    _Column2Items.Add(key, id);
                }
            }

            foreach (string item in _Column2Items.Keys)
            {
                this.Column2.Items.Add(item);
            }
        }

        private void ReportBuilding(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(Global.ReportName + " 產生完成");

            EnableForm(true);
            Document doc = (Document)e.Result;
            doc.MailMerge.DeleteFields();
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = Global.ReportName + ".doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void DataBuilding(object sender, DoWorkEventArgs e)
        {
            _BW.ReportProgress(0);
            SaveSetting();

            _BW.ReportProgress(10);
            //MappingData
            _MappingData = new Dictionary<string, List<string>>();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    string tagName = row.Cells[0].Value.ToString();

                    if (!_MappingData.ContainsKey(tagName))
                    {
                        _MappingData.Add(tagName, new List<string>());
                    }

                    if (row.Cells[1].Value != null)
                    {
                        string tagText = row.Cells[1].Value.ToString();
                        if (_Column2Items.ContainsKey(tagText))
                        {
                            string tagId = _Column2Items[tagText];
                            if (!_MappingData[tagName].Contains(tagId))
                            {
                                _MappingData[tagName].Add(tagId);
                            }
                        }
                    }
                }
            }

            Dictionary<string, StudentObj> studentDic = new Dictionary<string, StudentObj>();
            List<string> students = K12.Presentation.NLDPanels.Student.SelectedSource;
            string ids = string.Join("','", students);

            _BW.ReportProgress(20);
            //基本資料
            DataTable dt = _Q.Select("SELECT student.id,student.name,student.id_number,class.class_name,student.seat_no,student.student_number FROM student LEFT JOIN class ON ref_class_id = class.id WHERE student.id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                StudentObj obj = new StudentObj(row);
                if (!studentDic.ContainsKey(obj.Id))
                {
                    studentDic.Add(obj.Id, obj);
                }
            }

            //基本資料-TagId
            dt = _Q.Select("SELECT ref_student_id,ref_tag_id FROM tag_student WHERE ref_student_id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                string id = row["ref_student_id"].ToString();
                string tagid = row["ref_tag_id"].ToString();
                if (studentDic.ContainsKey(id))
                {
                    if (!studentDic[id].TagIds.Contains(tagid))
                    {
                        studentDic[id].TagIds.Add(tagid);
                    }
                }
            }

            _BW.ReportProgress(30);
            //服務學習紀錄
            dt = _Q.Select("SELECT ref_student_id,hours FROM $k12.service.learning.record WHERE ref_student_id IN ('" + ids + "')");
            foreach (DataRow row in dt.Rows)
            {
                string id = row["ref_student_id"].ToString();
                if (studentDic.ContainsKey(id))
                {
                    studentDic[id].ServiceHours += decimal.Parse(row["hours"].ToString());
                }
            }

            _BW.ReportProgress(40);
            //幹部紀錄
            dt = _Q.Select("SELECT studentid,schoolyear,semester FROM $behavior.thecadre WHERE studentid IN ('" + ids + "')");
            List<string> checkList = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                string id = row["studentid"].ToString();
                string schoolyear = row["schoolyear"].ToString();
                string semester = row["semester"].ToString();

                string key = id + "_" + schoolyear + "_" + semester;

                if (!checkList.Contains(key))
                {
                    if (studentDic.ContainsKey(id))
                    {
                        studentDic[id].CadreTimes++;
                        checkList.Add(key);
                    }
                }
            }

            _BW.ReportProgress(50);
            //獎懲紀錄
            List<AutoSummaryRecord> records = AutoSummary.Select(students, null);
            foreach (AutoSummaryRecord record in records)
            {
                string id = record.RefStudentID;
                if (studentDic.ContainsKey(id))
                {
                    studentDic[id].MeritA += record.MeritA;
                    studentDic[id].MeritB += record.MeritB;
                    studentDic[id].MeritC += record.MeritC;
                    studentDic[id].DemeritA += record.DemeritA;
                    studentDic[id].DemeritB += record.DemeritB;
                    studentDic[id].DemeritC += record.DemeritC;
                }
            }
            //List<DisciplineRecord> records = Discipline.SelectByStudentIDs(students);
            //foreach (DisciplineRecord record in records)
            //{
            //    string id = record.RefStudentID;
            //    if (studentDic.ContainsKey(id))
            //    {
            //        studentDic[id].MeritA += record.MeritA.HasValue ? record.MeritA.Value : 0;
            //        studentDic[id].MeritB += record.MeritB.HasValue ? record.MeritB.Value : 0;
            //        studentDic[id].MeritC += record.MeritC.HasValue ? record.MeritC.Value : 0;
            //        if (record.Cleared != "是")
            //        {
            //            studentDic[id].DemeritA += record.DemeritA.HasValue ? record.DemeritA.Value : 0;
            //            studentDic[id].DemeritB += record.DemeritB.HasValue ? record.DemeritB.Value : 0;
            //            studentDic[id].DemeritC += record.DemeritC.HasValue ? record.DemeritC.Value : 0;
            //        }
            //    }
            //}

            //取得功過換算比例
            MeritDemeritReduceRecord mdrr = MeritDemeritReduce.Select();
            int MAB = mdrr.MeritAToMeritB.HasValue ? mdrr.MeritAToMeritB.Value : 0;
            int MBC = mdrr.MeritBToMeritC.HasValue ? mdrr.MeritBToMeritC.Value : 0;
            int DAB = mdrr.DemeritAToDemeritB.HasValue ? mdrr.DemeritAToDemeritB.Value : 0;
            int DBC = mdrr.DemeritBToDemeritC.HasValue ? mdrr.DemeritBToDemeritC.Value : 0;

            //獎懲紀錄功過相抵
            foreach (StudentObj obj in studentDic.Values)
            {
                //if (!obj.HasDemeritAB)
                //{
                    obj.MeritC += ((obj.MeritA * MAB) + obj.MeritB) * MBC;
                    obj.DemeritC += ((obj.DemeritA * DAB) + obj.DemeritB) * DBC;

                    int total = obj.MeritC - obj.DemeritC;

                    if (total > 0)
                    {
                        obj.MC = total % MBC;
                        obj.MB = (total / MBC) % MAB;
                        obj.MA = (total / MBC) / MAB;
                    }
                    else if (total < 0)
                    {
                        total *= -1;
                        obj.DC = total % DBC;
                        obj.DB = (total / DBC) % DAB;
                        obj.DA = (total / DBC) / DAB;
                    }
                //}
            }

            _BW.ReportProgress(60);
            //體適能
            //先確認UDT存在
            dt = _Q.Select("SELECT name FROM _udt_table where name='ischool_student_fitness'");
            if (dt.Rows.Count > 0)
            {
                dt = _Q.Select("SELECT ref_student_id,sit_and_reach_degree, standing_long_jump_degree, sit_up_degree, cardiorespiratory_degree FROM $ischool_student_fitness WHERE ref_student_id IN ('" + ids + "')");
                foreach (DataRow row in dt.Rows)
                {
                    string id = row["ref_student_id"].ToString();
                    if (studentDic.ContainsKey(id))
                    {
                        //擇優判斷
                        if (GetScore(row) > studentDic[id].SportFitnessScore)
                        {
                            studentDic[id].sit_and_reach_degree = row["sit_and_reach_degree"].ToString();
                            studentDic[id].sit_up_degree = row["sit_up_degree"].ToString();
                            studentDic[id].standing_long_jump_degree = row["standing_long_jump_degree"].ToString();
                            studentDic[id].cardiorespiratory_degree = row["cardiorespiratory_degree"].ToString();
                        }
                    }
                }
            }

            _BW.ReportProgress(70);
            //均衡學習-年級對照
            Dictionary<string, Dictionary<string, string>> SchoolyearSemesteerToGrade = new Dictionary<string, Dictionary<string, string>>();
            foreach (JHSemesterHistoryRecord record in JHSemesterHistory.SelectByStudentIDs(students))
            {
                foreach (SemesterHistoryItem item in record.SemesterHistoryItems)
                {
                    if (!SchoolyearSemesteerToGrade.ContainsKey(item.RefStudentID))
                    {
                        SchoolyearSemesteerToGrade.Add(item.RefStudentID, new Dictionary<string, string>());
                    }

                    string key = item.SchoolYear + "_" + item.Semester;
                    if (!SchoolyearSemesteerToGrade[item.RefStudentID].ContainsKey(key))
                    {
                        if (item.Semester == 1)
                            SchoolyearSemesteerToGrade[item.RefStudentID].Add(key, item.GradeYear + "上");
                        else if (item.Semester == 2)
                            SchoolyearSemesteerToGrade[item.RefStudentID].Add(key, item.GradeYear + "下");
                    }

                }
            }

            _BW.ReportProgress(80);
            //均衡學習-領域分數
            List<JHSemesterScoreRecord> recs = JHSemesterScore.SelectByStudentIDs(students);
            foreach (JHSemesterScoreRecord rec in recs)
            {
                foreach (DomainScore score in rec.Domains.Values)
                {
                    string id = score.RefStudentID;
                    string key = score.SchoolYear + "_" + score.Semester;
                    string grade = "";

                    if (SchoolyearSemesteerToGrade.ContainsKey(id))
                    {
                        if (SchoolyearSemesteerToGrade[id].ContainsKey(key))
                        {
                            grade = SchoolyearSemesteerToGrade[id][key];
                        }
                    }

                    if (studentDic.ContainsKey(id))
                    {
                        string domain = score.Domain;

                        if ((domain == "健康與體育" || domain == "藝術與人文" || domain == "綜合活動") && !string.IsNullOrWhiteSpace(grade))
                        {
                            if (!studentDic[id].DomainScores.ContainsKey(domain))
                            {
                                studentDic[id].DomainScores.Add(domain, new Dictionary<string, decimal>());
                            }

                            if (!studentDic[id].DomainScores[domain].ContainsKey(grade))
                            {
                                decimal value = score.Score.HasValue ? score.Score.Value : 0;
                                studentDic[id].DomainScores[domain].Add(grade, value);
                            }
                        }
                    }
                }
            }

            //排序
            List<StudentObj> list = studentDic.Values.ToList();
            list.Sort(SortStudent);

            int progress = 80;
            decimal per = (decimal)(100-progress) / studentDic.Count;
            int count = 0;
            //Objects轉DataTable
            DataTable data = new DataTable();
            data.Columns.Add("學校名稱");
            data.Columns.Add("學校代碼");
            data.Columns.Add("班級");
            data.Columns.Add("姓名");
            data.Columns.Add("座號");
            data.Columns.Add("學號");
            data.Columns.Add("身分證字號");
            data.Columns.Add("服務時數");
            data.Columns.Add("幹部紀錄");
            data.Columns.Add("服務學習");
            data.Columns.Add("處分紀錄");
            data.Columns.Add("嘉獎");
            data.Columns.Add("小功");
            data.Columns.Add("大功");
            data.Columns.Add("警告");
            data.Columns.Add("小過");
            data.Columns.Add("大過");
            data.Columns.Add("表現評量");
            data.Columns.Add("坐姿體前彎");
            data.Columns.Add("立定跳遠");
            data.Columns.Add("仰臥起坐");
            data.Columns.Add("心肺適能");
            data.Columns.Add("體適能");
            data.Columns.Add("健康與體育");
            data.Columns.Add("藝術與人文");
            data.Columns.Add("綜合活動");
            data.Columns.Add("均衡學習");
            data.Columns.Add("對照身分");
            data.Columns.Add("弱勢身分");
            data.Columns.Add("弱勢身分_總");
            data.Columns.Add("均衡學習_總");
            data.Columns.Add("多元學習表現");
            foreach (StudentObj obj in list)
            {
                DataRow row = data.NewRow();
                row["學校名稱"] = _SchoolName;
                row["學校代碼"] = _SchoolCode;
                row["班級"] = obj.ClassName;
                row["姓名"] = obj.Name;
                row["座號"] = obj.SeatNo;
                row["學號"] = obj.StudentNumber;
                row["身分證字號"] = obj.IdNumber;
                row["服務時數"] = obj.ServiceHours;
                row["幹部紀錄"] = obj.CadreTimes;
                row["服務學習"] = obj.ServiceHoursScore;
                row["處分紀錄"] = obj.HasDemeritAB ? "有" : "無";
                row["嘉獎"] = obj.MC;
                row["小功"] = obj.MB;
                row["大功"] = obj.MA;
                row["警告"] = obj.DC;
                row["小過"] = obj.DB;
                row["大過"] = obj.DA;
                row["表現評量"] = obj.MeritDemeritScore;
                row["坐姿體前彎"] = obj.CheckScore("坐姿體前彎");
                row["立定跳遠"] = obj.CheckScore("立定跳遠");
                row["仰臥起坐"] = obj.CheckScore("仰臥起坐");
                row["心肺適能"] = obj.CheckScore("心肺適能");
                row["體適能"] = obj.SportFitnessScore;

                Dictionary<string, decimal> dic = obj.GetDomainScores();
                row["健康與體育"] = dic.ContainsKey("健康與體育") ? dic["健康與體育"] : 0;
                row["藝術與人文"] = dic.ContainsKey("藝術與人文") ? dic["藝術與人文"] : 0;
                row["綜合活動"] = dic.ContainsKey("綜合活動") ? dic["綜合活動"] : 0;
                row["均衡學習"] = obj.DomainItemScore;

                string[] tags = CheckTagId(obj.TagIds);
                row["對照身分"] = tags[0];
                row["弱勢身分"] = tags[1];

                row["弱勢身分_總"] = row["弱勢身分"].ToString();
                row["均衡學習_總"] = row["均衡學習"].ToString();
                row["多元學習表現"] = obj.ServiceHoursScore + obj.MeritDemeritScore + obj.SportFitnessScore;
                data.Rows.Add(row);

                count++;
                progress += (int)(count * per);
                _BW.ReportProgress(progress);
            }

            Document doc = _Configure.Template;
            doc.MailMerge.Execute(data);

            e.Result = doc;
        }

        private int SortStudent(StudentObj x, StudentObj y)
        {
            string xx = x.ClassName.PadLeft(20, '0');
            xx += x.SeatNo.PadLeft(3, '0');
            xx += x.StudentNumber.PadLeft(20, '0');

            string yy = y.ClassName.PadLeft(20, '0');
            yy += y.SeatNo.PadLeft(3, '0');
            yy += y.StudentNumber.PadLeft(20, '0');

            return xx.CompareTo(yy);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            ConfigureMaker();

            if (_BW.IsBusy)
            {
                MessageBox.Show("系統忙碌中,請稍後再試");
            }
            else
            {
                EnableForm(false);
                _BW.RunWorkerAsync();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ConfigureMaker();
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = Global.ReportName + "(範本).doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    Document doc = _Configure.Template;
                    doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("範本儲存失敗");
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //ConfigureMaker();
            List<Configure> Configures = _A.Select<Configure>();
            _A.DeletedValues(Configures);

            _Configure = new Configure();

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Document(dialog.FileName);
                    _Configure.Encode();
                    _Configure.Save();
                }
                catch
                {
                    MessageBox.Show("範本上傳失敗");
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("確認移除目前範本?", "ischool", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                List<Configure> Configures = _A.Select<Configure>();

                if (Configures.Count > 0)
                    _A.DeletedValues(Configures);

                _Configure = null;
            }
        }

        private void ConfigureMaker()
        {
            List<Configure> Configures = _A.Select<Configure>();

            if (Configures.Count > 0)
            {
                _Configure = Configures[0];
                _Configure.Decode();
            }
            else
            {
                _Configure = new Configure();
                TemplateSelecter selecter = new TemplateSelecter();
                selecter.ShowDialog();
                if (selecter.DialogResult == DialogResult.OK)
                {
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.Template_中區));
                }
                else
                {
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.Template_南區));
                }
                _Configure.Encode();
                _Configure.Save();
            }
        }

        private void EnableForm(bool enable)
        {
            this.buttonX1.Enabled = enable;
            this.linkLabel1.Enabled = enable;
            this.linkLabel2.Enabled = enable;
            this.linkLabel3.Enabled = enable;
            this.linkLabel4.Enabled = enable;
        }

        private int GetScore(DataRow row)
        {
            string sit_and_reach_degree = row["sit_and_reach_degree"].ToString();
            string standing_long_jump_degree = row["standing_long_jump_degree"].ToString();
            string sit_up_degree = row["sit_up_degree"].ToString();
            string cardiorespiratory_degree = row["cardiorespiratory_degree"].ToString();

            int score = 0;

            //sit_and_reach_degree
            if (sit_and_reach_degree == "金牌" || sit_and_reach_degree == "銀牌" || sit_and_reach_degree == "銅牌" || sit_and_reach_degree == "中等")
            {
                score += 2;
            }
            //standing_long_jump_degree
            if (standing_long_jump_degree == "金牌" || standing_long_jump_degree == "銀牌" || standing_long_jump_degree == "銅牌" || standing_long_jump_degree == "中等")
            {
                score += 2;
            }
            //sit_up_degree
            if (sit_up_degree == "金牌" || sit_up_degree == "銀牌" || sit_up_degree == "銅牌" || sit_up_degree == "中等")
            {
                score += 2;
            }
            //cardiorespiratory_degree
            if (cardiorespiratory_degree == "金牌" || cardiorespiratory_degree == "銀牌" || cardiorespiratory_degree == "銅牌" || cardiorespiratory_degree == "中等")
            {
                score += 2;
            }

            if (score > 6)
                score = 6;

            return score;
        }

        private string[] CheckTagId(List<string> list)
        {
            List<string> retVal = new List<string>();

            foreach (string tagName in _MappingData.Keys)
            {
                foreach (string tagId in _MappingData[tagName])
                {
                    if (list.Contains(tagId))
                    {
                        if (!retVal.Contains(tagName))
                        {
                            retVal.Add(tagName);
                        }
                    }
                }
            }

            string[] str = new string[2];
            if (retVal.Count > 0)
            {
                str[0] = string.Join("、", retVal);
                str[1] = "2";
            }
            else
            {
                str[0] = "   ";
                str[1] = "0";
            }

            return str;
        }

        private void SaveSetting()
        {
            List<Setting> UDTlist = _A.Select<Setting>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist.Clear(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value == null) //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                Setting obj = new Setting();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }

        private void LoadSetting()
        {
            List<Setting> UDTlist = _A.Select<Setting>(); //檢查UDT並回傳資料
            DataGridViewRow row;
            if (UDTlist.Count > 0) //UDT內有設定才做讀取
            {
                for (int i = 0; i < UDTlist.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = UDTlist[i].Target;
                    row.Cells[1].Value = UDTlist[i].Source;
                    dataGridViewX1.Rows.Add(row);
                }
            }
        }

        private void Report_Load(object sender, EventArgs e)
        {
            LoadSetting();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = Global.ReportName + "(合併欄位).doc";
            sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    Document doc = new Document(new MemoryStream(Properties.Resources.合併欄位));
                    doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }
    }
}
