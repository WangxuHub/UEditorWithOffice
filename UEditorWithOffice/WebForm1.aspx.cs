using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UEditorWithOffice
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        //上传按钮
        protected void Button1_Click(object sender, EventArgs e)
        {
            var file = FileUpload1.PostedFile;

            var ext = System.IO.Path.GetExtension(FileUpload1.FileName).ToLower();

            string conventUrl = "";
            switch (ext)
            {
                case ".doc":
                case ".docx":
                    conventUrl = WordConvent(file);
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
                conventUrl = conventUrl.Replace(MapPath("/"), "").Replace("\\",@"/");
                Response.Write("<script>window.open('" + conventUrl + "')</script>");
            }
            else
            {
                Response.Write("<script>alert('无效的格式')</script>");
            }
        }

        private string WordConvent(HttpPostedFile file)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(file.InputStream);

            var fileName = GetSaveHtmlName();

            doc.Save(fileName,Aspose.Words.SaveFormat.Html);

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
            wk.Save(fileName,Aspose.Cells.SaveFormat.Html);
            return fileName;
        }

        private string PDFConvent(HttpPostedFile file)
        {
            Aspose.Pdf.Document pdf = new Aspose.Pdf.Document(file.InputStream);
            var fileName = GetSaveHtmlName();
            pdf.Save(fileName,Aspose.Pdf.SaveFormat.Html);
            return fileName;
        }


        private string GetSaveHtmlName()
        {
            string fileName = Guid.NewGuid() + ".html";
            string filePath = "/temp/" + fileName;

            string absolutePath = MapPath(filePath);

            string dir = System.IO.Path.GetDirectoryName(absolutePath);
            if(!System.IO.Directory.Exists(dir))
            {
                System.IO.Directory.CreateDirectory(dir);
            }

            return absolutePath;
        }
    }
}