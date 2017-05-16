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

        mJsonresult json = new mJsonresult();
        protected override void OnPreInit(EventArgs e)
        {
            string postType = Request["PostType"];
            if (string.IsNullOrWhiteSpace(postType))
            {
                base.OnPreInit(e);
                return;
            }

            try
            {
                switch (postType)
                {
                    case "Submit":
                        Submit();
                        break;
                    case "Load":
                        LoadContent();
                        break;
                }
            }
            catch(Exception ee)
            {
                json = new mJsonresult();
                json.msg = ee.Message;
                json.success = false;
                
            }

            string res = "";
            if (json != null)
            {
                res = Newtonsoft.Json.JsonConvert.SerializeObject(json);
            }

            Response.Write(res);
            Response.End();

        }

        public static string cacheString = "";

        private void Submit()
        {
            cacheString = HttpUtility.UrlDecode(Request["content"]??"");
            json.success = true;
        }

        private void LoadContent()
        {
            json.obj = cacheString;
            json.success = true;
        }
    }

    public class mJsonresult
    {
        public bool success;

        public object obj;

        public string msg;
    }
}