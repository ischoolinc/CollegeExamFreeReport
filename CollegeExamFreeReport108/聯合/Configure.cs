using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.UDT;

namespace CollegeExamFreeReport108
{
    [FISCA.UDT.TableName(Global._UDTTableName)]
    class Configure : ActiveRecord
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

        ///<summary>
        /// 檢查是否上傳電子報表
        ///</summary>
        [Field(Field = "check_upload_epaper", Indexed = false)]
        public bool CheckUploadEpaper { get; set; }
    }
}
