using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport108
{
    class StudentObj
    {
        public string Id, Name, ClassName, IdNumber, SeatNo, StudentNumber;
        /// <summary>
        /// 服務學習總時數
        /// </summary>
        public decimal ServiceHours;

        //原始獎懲紀錄
        public int MeritA, MeritB, MeritC, DemeritA, DemeritB, DemeritC;
        //功過相抵後獎懲紀錄
        public int MA, MB, MC, DA, DB, DC;
        //體適能常模
        public string sit_and_reach_degree, standing_long_jump_degree, sit_up_degree, cardiorespiratory_degree;

        public Dictionary<string, Dictionary<string, decimal>> DomainScores;
        private Dictionary<string, decimal> domainAverageScores;

        public List<string> TagIds;

        public Dictionary<string, decimal> GetDomainScores()
        {
            domainAverageScores.Clear();

            foreach (string domain in DomainScores.Keys)
            {
                int count = 0;

                if (!domainAverageScores.ContainsKey(domain))
                {
                    domainAverageScores.Add(domain, 0);
                }

                foreach (KeyValuePair<string, decimal> kvp in DomainScores[domain])
                {
                    if (kvp.Key == "1上" || kvp.Key == "1下" || kvp.Key == "2上" || kvp.Key == "2下" || kvp.Key == "3上" || kvp.Key == "7上" || kvp.Key == "7下" || kvp.Key == "8上" || kvp.Key == "8下" || kvp.Key == "9上")
                    {
                        domainAverageScores[domain] += kvp.Value;
                        count++;
                    }
                }

                if (count > 0)
                    domainAverageScores[domain] /= count;

                domainAverageScores[domain] = Math.Round(domainAverageScores[domain], 2, MidpointRounding.AwayFromZero);
            }

            return domainAverageScores;
        }


        /// <summary>
        /// 聯合均衡學習：每一個領域滿60分得2分，上限6分。(2021-12 Cynthia)
        /// </summary>
        /// https://3.basecamp.com/4399967/buckets/15852426/todos/6692639656   2023/10/26
        public int DomainItemScore
        {
            get
            {
                int score = 0;

                if (domainAverageScores.ContainsKey("健康與體育"))
                {
                    if (domainAverageScores["健康與體育"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (domainAverageScores.ContainsKey("藝術"))
                {
                    if (domainAverageScores["藝術"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (domainAverageScores.ContainsKey("科技"))
                {
                    if (domainAverageScores["科技"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (domainAverageScores.ContainsKey("綜合活動"))
                {
                    if (domainAverageScores["綜合活動"] >= 60)
                    {
                        score += 2;
                    }
                }

                if (score > 6)
                    score = 6;

                return score;
            }
        }

        /// <summary>
        /// 優先均衡學習：每一個領域滿60分得7分，上限21分。(2021-12 Cynthia)
        /// </summary>
        public int DomainItemScore_Priority
        {
            get
            {
                int score = 0;

                if (domainAverageScores.ContainsKey("健康與體育"))
                {
                    if (domainAverageScores["健康與體育"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("科技"))
                {
                    if (domainAverageScores["科技"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("藝術"))
                {
                    if (domainAverageScores["藝術"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("綜合活動"))
                {
                    if (domainAverageScores["綜合活動"] >= 60)
                    {
                        score += 7;
                    }
                }
                if (score > 21)
                    score = 21;

                return score;
            }
        }

        /// <summary>
        /// 完全免試 均衡學習：每一個領域滿60分得7分，上限28分。(2024-4 Dylan)
        /// </summary>
        public int DomainItemScore_Completely
        {
            get
            {
                int score = 0;

                if (domainAverageScores.ContainsKey("健康與體育"))
                {
                    if (domainAverageScores["健康與體育"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("科技"))
                {
                    if (domainAverageScores["科技"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("藝術"))
                {
                    if (domainAverageScores["藝術"] >= 60)
                    {
                        score += 7;
                    }
                }

                if (domainAverageScores.ContainsKey("綜合活動"))
                {
                    if (domainAverageScores["綜合活動"] >= 60)
                    {
                        score += 7;
                    }
                }
                if (score > 28)
                    score = 28;

                return score;
            }
        }

        /// <summary>
        /// 聯合服務時數：8小時1分，上限7分,old:4小時1分，上限7分
        /// </summary>
        /// https://3.basecamp.com/4399967/buckets/15852426/todos/6692639656   2023/10/26
        public int ServiceHoursScore
        {
            get
            {
                //int score = (int)(ServiceHours / 8);
                //score += CadreTimes;
                int score = (int)(ServiceHours / 8);
                score += CadreTimes;

                if (score > 7)
                    score = 7;

                return score;
            }
        }

        /// <summary>
        /// 優先服務學習: 每1小時0.25分，上限15分，old:每1小時0.5分，上限15分。
        /// </summary>
        /// https://3.basecamp.com/4399967/buckets/15852426/todos/4417454620
        public int ServiceHoursScore_Priority
        {
            get
            {
                //double score = (double)(ServiceHours) *0.25;
                //score += CadreTimes*2;
                decimal score = (int)(ServiceHours) *0.25m;
                score += CadreTimes * 2;

                if (score > 15)
                    score = 15;

                return (int)Math.Floor(score);
            }
        }

        /// <summary>
        /// 完全免試 服務學習: 每1小時0.5分，上限15分 (2024-8 Dylan)
        /// </summary>
        /// https://3.basecamp.com/4399967/buckets/15852426/todos/7221323792
        public int ServiceHoursScore_Completely
        {
            get
            {
                decimal score = (int)(ServiceHours) *0.5m; //服務學習

                if (score > 15)
                    score = 15;

                return (int)Math.Floor(score);
            }
        }

        /// <summary>
        /// 幹部學期數
        /// </summary>
        public int CadreTimes;

        /// <summary>
        /// 幹部積分
        /// </summary>
        public decimal CadreTimes_Completely
        {
            get
            {
                decimal score = CadreTimes * 2; //幹部紀錄

                return score;
            }
        }

        public int MeritDemeritScore
        {
            get
            {
                int score = 0;

                if (DA == 0 && DB == 0 && DC == 0)
                {
                    score = 1;
                }

                if (!HasDemeritAB)
                {
                    if (MA > 0)
                    {
                        score = 4;
                    }
                    else if (MB > 0)
                    {
                        score = 3;
                    }
                    else if (MC > 0)
                    {
                        score = 2;
                    }
                }

                return score;
            }
        }

        public bool HasDemeritAB
        {
            get
            {
                if (DemeritA == 0 && DemeritB == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        public int SportFitnessScore
        {
            get
            {
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

                // 免測處理
                if (sit_and_reach_degree == "免測" || standing_long_jump_degree == "免測" || sit_up_degree == "免測" || cardiorespiratory_degree == "免測")
                {
                    score = 6;
                }

                return score;
            }
        }

        public string CheckScore(string str)
        {
            string value = "";
            switch (str)
            {
                case "坐姿體前彎":
                    value = sit_and_reach_degree;
                    break;

                case "立定跳遠":
                    value = standing_long_jump_degree;
                    break;

                case "仰臥起坐":
                    value = sit_up_degree;
                    break;

                case "心肺適能":
                    value = cardiorespiratory_degree;
                    break;

                default:
                    value = "";
                    break;
            }

            if (value == "金牌" || value == "銀牌" || value == "銅牌" || value == "中等" || value=="免測")
            {
                return "達";
            }            
            else
            {
                return "未達";
            }
        }

        public void MeritDemeritTransfer()
        {
            int merit = ((MeritA * Report.MAB) + MeritB) * Report.MBC + MeritC;
            int demerit = ((DemeritA * Report.DAB) + DemeritB) * Report.DBC + DemeritC;

            int total = merit - demerit;

            if (total > 0)
            {
                MC = total % Report.MBC;
                MB = (total / Report.MBC) % Report.MAB;
                MA = (total / Report.MBC) / Report.MAB;
                /*
                //最小單位先存起來
                MC = total;

                //原始紀錄有大功或小功必須先轉換一次
                if (MeritA > 0 || MeritB > 0)
                {
                    MB = MC / Report.MBC;
                    MC = MC % Report.MBC;
                }

                //原始紀錄有大功再轉一次
                if (MeritA > 0)
                {
                    MA = MB / Report.MAB;
                    MB = MB % Report.MAB;
                }
                 * */
            }
            else if (total < 0)
            {
                total *= -1;
                DC = total % Report.DBC;
                DB = (total / Report.DBC) % Report.DAB;
                DA = (total / Report.DBC) / Report.DAB;
            }
        }



        public void MeritDemeritTransfer_priority()
        {
            int merit = ((MeritA * Report_priority.MAB) + MeritB) * Report_priority.MBC + MeritC;
            int demerit = ((DemeritA * Report_priority.DAB) + DemeritB) * Report_priority.DBC + DemeritC;

            int total = merit - demerit;

            if (total > 0)
            {
                MC = total % Report_priority.MBC;
                MB = (total / Report_priority.MBC) % Report_priority.MAB;
                MA = (total / Report_priority.MBC) / Report_priority.MAB;
                /*
                //最小單位先存起來
                MC = total;

                //原始紀錄有大功或小功必須先轉換一次
                if (MeritA > 0 || MeritB > 0)
                {
                    MB = MC / Report.MBC;
                    MC = MC % Report.MBC;
                }

                //原始紀錄有大功再轉一次
                if (MeritA > 0)
                {
                    MA = MB / Report.MAB;
                    MB = MB % Report.MAB;
                }
                 * */
            }
            else if (total < 0)
            {
                total *= -1;
                DC = total % Report_priority.DBC;
                DB = (total / Report_priority.DBC) % Report_priority.DAB;
                DA = (total / Report_priority.DBC) / Report_priority.DAB;
            }
        }

        public void MeritDemeritTransfer_Completely()
        {
            int merit = ((MeritA * Report_Completely.MAB) + MeritB) * Report_Completely.MBC + MeritC;
            int demerit = ((DemeritA * Report_Completely.DAB) + DemeritB) * Report_Completely.DBC + DemeritC;

            int total = merit - demerit;

            if (total > 0)
            {
                MC = total % Report_Completely.MBC;
                MB = (total / Report_Completely.MBC) % Report_Completely.MAB;
                MA = (total / Report_Completely.MBC) / Report_Completely.MAB;
                /*
                //最小單位先存起來
                MC = total;

                //原始紀錄有大功或小功必須先轉換一次
                if (MeritA > 0 || MeritB > 0)
                {
                    MB = MC / Report.MBC;
                    MC = MC % Report.MBC;
                }

                //原始紀錄有大功再轉一次
                if (MeritA > 0)
                {
                    MA = MB / Report.MAB;
                    MB = MB % Report.MAB;
                }
                 * */
            }
            else if (total < 0)
            {
                total *= -1;
                DC = total % Report_Completely.DBC;
                DB = (total / Report_Completely.DBC) % Report_Completely.DAB;
                DA = (total / Report_Completely.DBC) / Report_Completely.DAB;
            }
        }
        public StudentObj(DataRow row)
        {
            this.Id = row["id"].ToString();
            this.Name = row["name"].ToString();
            this.ClassName = row["class_name"].ToString();
            this.IdNumber = row["id_number"].ToString();
            this.SeatNo = row["seat_no"].ToString();
            this.StudentNumber = row["student_number"].ToString();
            this.ServiceHours = 0;
            this.CadreTimes = 0;
            this.MeritA = 0;
            this.MeritB = 0;
            this.MeritC = 0;
            this.DemeritA = 0;
            this.DemeritB = 0;
            this.DemeritC = 0;
            this.MA = 0;
            this.MB = 0;
            this.MC = 0;
            this.DA = 0;
            this.DB = 0;
            this.DC = 0;
            sit_and_reach_degree = "";
            standing_long_jump_degree = "";
            sit_up_degree = "";
            cardiorespiratory_degree = "";
            DomainScores = new Dictionary<string, Dictionary<string, decimal>>();
            domainAverageScores = new Dictionary<string, decimal>();
            TagIds = new List<string>();
        }
    }
}
