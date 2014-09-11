using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.IO;
using System.Net.Mime;

namespace Brilliantech.Export.Report
{
    public class ReportExporter
    {
        private HttpContext context;
        private string filePath;
        private string xml;

        public ReportExporter() { }
        public ReportExporter(HttpContext context, string filePath,string xml) {
            this.context = context;
            this.filePath = filePath;
            this.xml = xml;
        }
        public void ProcessExportTableChartRequest()
        {
            IReport report = new ReportTableChart(filePath, xml);
            report.Generate();
            ResponseReport();
        }

        private void ResponseReport()
        {
            HttpResponse resp = this.context.Response;

            resp.ContentType = this.ContentType;
            resp.HeaderEncoding = Encoding.UTF8;
            resp.AppendHeader("Content-Disposition", "attachment;filename=my_report.xlsx");
            resp.AppendHeader("Cache-Control", "max-age=0");

            MemoryStream ms = new MemoryStream();
            using (FileStream excel = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                excel.CopyTo(ms);
            }
            File.Delete(filePath);
            ms.WriteTo(resp.OutputStream);
        }
        public string ContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }
    }
}
