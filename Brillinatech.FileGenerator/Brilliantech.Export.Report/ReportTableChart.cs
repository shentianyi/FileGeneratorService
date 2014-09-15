using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using Brilliantech.Export.Report.XmlParser;
using Brilliantech.Export.Report.Chart;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;

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
                var chart = worksheet.Drawings.AddChart(charts[i].Title, eChartType.ColumnStacked) as ExcelChart;
                if (!charts[i].ShowLegend)
                {
                    chart.Legend.Remove();
                }
                chart.Title.Text = charts[i].Title;
                chart.SetPosition(Convert.ToInt32(chartOffsetHeight), 10);
                if (charts[i].Height.HasValue && charts[i].Width.HasValue)
                {
                    chart.SetSize(charts[i].Width.Value, charts[i].Height.Value);
                }
                RChartType[] chartTypes=charts[i].ChartTypes;

                for (var n = 0; n < chartTypes.Length; n++)
                {

                    ExcelChart chartType = chart.PlotArea.ChartTypes.Add(chartTypes[n].Type);
                    RSerie[] series = chartTypes[n].Series;
                    for (var j = 0; j < series.Length; j++)
                    {
                        var cc = GetChartType(chartTypes[n].Type, chartType, series[j]);
                        ExcelChartSerie serie = cc.Series.Add(worksheet.Cells[series[j].YAixs], worksheet.Cells[series[j].XAixs]);

                        serie = GetChartSerie(chartTypes[n].Type, serie, series[j]);

                        if (series[j].HeaderAddress != null)
                        {
                            serie.HeaderAddress = worksheet.Cells[series[j].HeaderAddress];
                        }
                        if (series[j].UseSecondaryAxis)
                        {
                            chartType.UseSecondaryAxis = true;
                            chartType.YAxis.Format = series[j].YAxisFormatString;
                        }
                    }
                    chartOffsetHeight += (double)charts[i].Height.Value + 10;
                }
                //  chartOffsetHeight += chart.To.RowOff * defaultRowHeight/0.75+10; 
            }
        }

        private dynamic GetChartType(eChartType type, ExcelChart chart, RSerie serie)
        {
            if (type.Equals(eChartType.Line))
            {
                var c = chart as ExcelLineChart;
                if (serie.ShowDataLabel)
                    c.DataLabel.ShowValue = true;
                c.DataLabel.ShowPercent = true;
                return c;
            }
            else if (type.Equals(eChartType.ColumnStacked))
            {
                var c = chart as ExcelBarChart;
                if (serie.ShowDataLabel)
                    c.DataLabel.ShowValue = true;

                return c;
            }
            var cc = chart as ExcelBarChart;
            return cc;
        }

        private dynamic GetChartSerie(eChartType type,ExcelChartSerie serie,RSerie rserie) {

            if (type.Equals(eChartType.Line))
            {
                var s = serie as ExcelLineChartSerie;
                if (rserie.ColorString != null)
                {
                    if (rserie.ColorString != null)
                        s.LineColor = rserie.ColorString;
                    // not use, i cannot change the marker color!
                    // s.Marker = eMarkerStyle.Diamond;     
                }
                return s;
            }
            else if (type.Equals(eChartType.ColumnStacked))
            {
                //var c = chart as ExcelBarChart;
                //if (serie.ShowDataLabel)
                //    c.DataLabel.ShowValue = true;
                //return c;
            } 
            return serie;
        }
    }
}
