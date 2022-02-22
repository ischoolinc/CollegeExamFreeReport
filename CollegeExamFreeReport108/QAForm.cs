﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace CollegeExamFreeReport108
{
    public partial class QAForm : BaseForm
    {
        public QAForm()
        {
            InitializeComponent();

            // InitDataGridView



            DataGridViewRow dgvrow2 = new DataGridViewRow();
            dgvrow2.CreateCells(dataGridViewX1);
            dgvrow2.Cells[0].Value = "多元學習表現_服務學習";
            dgvrow2.Cells[1].Value = "系統根據學生的學期歷程資料讀取學生就學期間的學年度、學期，進一步的去讀取學生的社團資料、幹部紀錄與服務學習紀錄。而需特別注意的部分是比序積分採計的社團資料為社團學期結算後資料，而服務學習紀錄根據採計截止日期的設定去計算。";
            dgvrow2.Cells[2].Value = "7分";
            dataGridViewX1.Rows.Add(dgvrow2);

            DataGridViewRow dgvrow3 = new DataGridViewRow();
            dgvrow3.CreateCells(dataGridViewX1);
            dgvrow3.Cells[0].Value = "多元學習表現_日常生活表現評量";
            dgvrow3.Cells[1].Value = "根據採計截止日期的設定去計算懲戒明細紀錄、懲戒非明細資料與銷過紀錄。";
            dgvrow3.Cells[2].Value = "4分";
            dataGridViewX1.Rows.Add(dgvrow3);

            DataGridViewRow dgvrow4 = new DataGridViewRow();
            dgvrow4.CreateCells(dataGridViewX1);
            dgvrow4.Cells[0].Value = "多元學習表現_體適能";
            dgvrow4.Cells[1].Value = "根據學生過往學期所有的體適能紀錄，擇優達門檻採計";
            dgvrow4.Cells[2].Value = "6分";
            dataGridViewX1.Rows.Add(dgvrow4);

            DataGridViewRow dgvrow5 = new DataGridViewRow();
            dgvrow5.CreateCells(dataGridViewX1);
            dgvrow5.Cells[0].Value = "弱勢身分";
            dgvrow5.Cells[1].Value = "根據設定對應系統類別低收入戶、中低收入 戶、直系血親尊親屬支領失業給付及特殊境遇家庭子女身分者";
            dgvrow5.Cells[2].Value = "2分";
            dataGridViewX1.Rows.Add(dgvrow5);

            DataGridViewRow dgvrow1 = new DataGridViewRow();
            dgvrow1.CreateCells(dataGridViewX1);
            dgvrow1.Cells[0].Value = "均衡學習";
            dgvrow1.Cells[1].Value = "系統根據學生的學期歷程資料讀取學生就學期間的學年度、學期，進一步的去比對學生的學期科目成績，將屬於健體、藝文、綜合這三個領域的學期科目成績做五學期的平均計算。";
            dgvrow1.Cells[2].Value = "6分";
            dataGridViewX1.Rows.Add(dgvrow1);

        }
    }
}
