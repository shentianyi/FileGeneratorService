using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using Brilliantech.Export.Report.XmlParser;
using Brilliantech.Export.Report.Chart;
using OfficeOpenXml.Drawing.Chart;

namespace Brilliantech.Export.Report
{
    public class ReportTableChart : Report, IReport
    {
        protected ChartXmlParser chartXmlParser;

        public ReportTableChart(string filePath, string xml)
        {
            this.filePath = filePath;
            this.fileInfo = new FileInfo(filePath);
            this.xml = xml;
            tableXmlParser = new TableXmlParser(xml);
            chartXmlParser = new ChartXmlParser(xml);
        }

        public void Generate()
        {
            try
            {
                using (package = new ExcelPackage(this.FileInfo))
                {
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    GenerateTableHead();
                    MergeTableHead();
                    GenerateTableBody();
                    GenerateCharts();
                    package.Save();
                }
            }
            catch (Exception e)
            {

            }
        }

        public void GenerateCharts()
        {
            double chartOffsetHeight = ((rowOffsetCount + 1) * defaultRowHeight) / 0.75;

            RChart[] charts = chartXmlParser.GetCharts();
            for (int i = 0; i < charts.Length; i++)
            {
                ExcelChart chart = null;
                chart = worksheet.Drawings.AddChart(charts[i].Title, eChartType.ColumnStacked) as ExcelChart;
                chart.Legend.Remove();
                chart.Title.Text = charts[i].Title;
                chart.SetPosition(Convert.ToInt32(chartOffsetHeight), 10);
                chart.SetSize(charts[i].Width, charts[i].Height);


                RSerie[] series = charts[i].Series;
                for (var j = 0; j < series.Length; j++)
                {
                    ExcelChart chartType = chart.PlotArea.ChartTypes.Add(series[j].Type);

                    GetChartType(series[j].Type,chartType).Series.Add(
                          worksheet.Cells[series[j].YStartRow, series[j].YStartCol, series[j].YEndRow, series[j].YEndCol], worksheet.Cells[series[j].XStartRow, series[j].XStartCol, series[j].XEndRow, series[j].XEndCol]);

                }
                chartOffsetHeight += charts[i].Height + 10;
            }
        }

        public dynamic GetChartType(eChartType type, ExcelChart chart)
        {
            if (type.Equals(eChartType.Line))
            {
                var c = chart as ExcelLineChart;
                c.DataLabel.ShowValue = true;
                return c;
            }
            var cc = chart as ExcelBarChart;
            cc.DataLabel.ShowValue = true;
            return cc;

        }
    }
}
