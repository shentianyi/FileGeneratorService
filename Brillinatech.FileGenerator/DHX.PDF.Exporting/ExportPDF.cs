using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using DHTMLX.Export.PDF;

namespace DHX.PDF.Exporting
{
  public  class ExportPDF
  {
      public static void ProcessExportRequest(HttpContext context)
      {
          var writer = new PDFWriter();
          var xml = context.Request.Form["grid_xml"];
          xml = context.Server.UrlDecode(xml);
          writer.Generate(xml, context.Response);

      }
    }
}
