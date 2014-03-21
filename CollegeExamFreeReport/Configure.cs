using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollegeExamFreeReport
{
    [FISCA.UDT.TableName(Global._UDTTableName)]
    class Configure : FISCA.UDT.ActiveRecord
    {
        /// <summary>
        /// 報表範本
        /// </summary>
        [FISCA.UDT.Field]
        private string TemplateStream { get; set; }
        public Aspose.Words.Document Template { get; set; }

        public void Encode()
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            this.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
        }

        public void Decode()
        {
            this.Template = new Aspose.Words.Document(new MemoryStream(Convert.FromBase64String(this.TemplateStream)));
        }
    }
}
