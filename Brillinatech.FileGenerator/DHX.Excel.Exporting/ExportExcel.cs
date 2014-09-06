using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using DHTMLX.Export.Excel;

namespace DHX.Excel.Exporting
{
    public class ExportExcel
    {
        public static void ProcessExportRequest(HttpContext context)
        {
            ExcelWriter writer = new ExcelWriter();
            var xml = context.Request.Form["grid_xml"];
            xml = context.Server.UrlDecode(xml);
            writer.Generate(xml, context.Response);
        }
    }
}
