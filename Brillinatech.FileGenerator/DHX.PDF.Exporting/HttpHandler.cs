using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace DHX.PDF.Exporting
{
    public class HttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            ExportPDF.ProcessExportRequest(context);
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
