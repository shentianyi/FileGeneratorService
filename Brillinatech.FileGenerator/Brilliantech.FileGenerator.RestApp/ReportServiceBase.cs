using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;

namespace Brilliantech.FileGenerator.RestApp
{
    public class ReportServiceBase
    {

        public static string ReportDefaultPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTmp", Guid.NewGuid().ToString() + ".xlsx");
        }

        public static string GetXmlString(HttpContext context)
        {
            var xml = context.Request.Form["grid_xml"];
            return context.Server.UrlDecode(xml);
        }
    }
}