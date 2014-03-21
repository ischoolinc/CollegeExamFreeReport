using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CollegeExamFreeReport
{
    public partial class TemplateSelecter : BaseForm
    {
        public TemplateSelecter()
        {
            InitializeComponent();
            this.ControlBox = false;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
