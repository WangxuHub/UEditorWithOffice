using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace UEditorWithOffice
{
    /// <summary>
    /// Upload 的摘要说明
    /// </summary>
    public class Upload : IHttpHandler
    {
        private HttpContext con = null;

        private mJsonResult json = new mJsonResult();
        public void ProcessRequest(HttpContext context)
        {
            con = context;

            string PostType = con.Request["PostType"];

            if(string.IsNullOrEmpty(PostType))
            {
                throw new ArgumentNullException("PostType");
            }

            switch(PostType)
            {
                case "UploadOffice":
                    UploadOffice();
                    break;
            }

            if(json!=null)
            {
                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(json);
                con.Response.Write(jsonStr);
                con.Response.End();
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }



        //上传按钮
        protected void UploadOffice()
        {
            var file = con.Request.Files[0];
            var conventPageType = con.Request["conventPageType"];

            var ext = System.IO.Path.GetExtension(file.FileName).ToLower();

            string conventUrl = "";
            switch (ext)
            {
                case ".doc":
                case ".docx":
                    conventUrl = WordConvent(file, conventPageType);
                    break;
                case ".ppt":
                case ".pptx":
                    conventUrl = PPTConvent(file);
                    break;
                case ".xls":
                case ".xlsx":
                    conventUrl = ExcelConvent(file);
                    break;
                case ".pdf":
                    conventUrl = PDFConvent(file);
                    break;
            }

            if (!string.IsNullOrEmpty(conventUrl))
            {
                conventUrl = "/"+conventUrl.Replace(con.Server.MapPath("/"), "").Replace("\\", @"/");
                json.success = true;
                json.obj = conventUrl;
            }
            else
            {
                json.msg = "无效的格式";
            }

        }

        private string WordConvent(HttpPostedFile file,string conventPageType)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(file.InputStream);

            var fileName = GetSaveHtmlName();

            if (conventPageType == "web")
            {
                doc.Save(fileName, Aspose.Words.SaveFormat.Html);
            }
            else if(conventPageType == "page")
            {
                doc.Save(fileName, Aspose.Words.SaveFormat.HtmlFixed);
            }
            return fileName;
        }

        private string PPTConvent(HttpPostedFile file)
        {
            Aspose.Slides.Presentation ppt = new Aspose.Slides.Presentation(file.InputStream);

            var fileName = GetSaveHtmlName();
            ppt.Save(fileName, Aspose.Slides.Export.SaveFormat.Html);
            return fileName;
        }

        private string ExcelConvent(HttpPostedFile file)
        {
            Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(file.InputStream);

            var fileName = GetSaveHtmlName();
            wk.Save(fileName, Aspose.Cells.SaveFormat.Html);
            return fileName;
        }

        private string PDFConvent(HttpPostedFile file)
        {
            Aspose.Pdf.Document pdf = new Aspose.Pdf.Document(file.InputStream);
            var fileName = GetSaveHtmlName();
            pdf.Save(fileName, Aspose.Pdf.SaveFormat.Html);
            return fileName;
        }


        private string GetSaveHtmlName()
        {
            string fileName = Guid.NewGuid() + ".html";
            string filePath = "/temp/" + fileName;

            string absolutePath = con.Server.MapPath(filePath);

            string dir = System.IO.Path.GetDirectoryName(absolutePath);
            if (!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            return absolutePath;
        }
    }

    public class mJsonResult
    {
        public int total;
        public bool success;
        public System.Collections.IEnumerable rows;
        public object obj;
        public string msg;
    }
}