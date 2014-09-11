using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Brilliantech.Export.Report.XmlParser;
using OfficeOpenXml;
using System.IO;

namespace Brilliantech.Export.Report
{
    public class ReportTable : Report, IReport
    {
        public ReportTable(string filePath, string xml)
        {
            this.filePath = filePath;
            this.fileInfo = new FileInfo(filePath);
            this.xml = xml;
            tableXmlParser = new TableXmlParser(xml);
            this.colCount = tableXmlParser.GetColCount();
        }

        public void Generate()
        {
            using (package = new ExcelPackage(this.FileInfo))
            {
                worksheet = package.Workbook.Worksheets.Add("Sheet1");
                GenerateTableHead();
                MergeTableHead();
                GenerateTableBody();
                package.Save();
            }
        }
    }
}
