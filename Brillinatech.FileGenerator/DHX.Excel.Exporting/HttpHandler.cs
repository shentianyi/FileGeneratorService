using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using DHTMLX.Export.Excel;

namespace DHX.Excel.Exporting
{
    public class HttpHandler: IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            ExportExcel.ProcessExportRequest(context);
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}
