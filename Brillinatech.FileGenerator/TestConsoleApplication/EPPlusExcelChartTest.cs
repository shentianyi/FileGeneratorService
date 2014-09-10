using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.Data;
using System.IO;
using System.Drawing;  
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;  

namespace TestConsoleApplication
{
   public class EPPlusExcelChartTest
    {
        private static readonly string[] MonthNames = new string[] { "一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月" };
        
        private static readonly string[] CommpanyNames = new string[] { "Microsoft","HH" };

        public static void Run() {
            string fileName = "ExcelReport-" + DateTime.Now.ToString("yyyy_MM_dd_HHmmss") + ".xlsx";
            string reportTitle = "2013年度五大公司实际情况与原计划的百分比";
            FileInfo file = new FileInfo("C:\\Excel\\" + fileName);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = null;
                
                ExcelChartSerie chartSerie = null;
                ExcelBarChart chart = null;
                #region research
                worksheet = package.Workbook.Worksheets.Add("Data");
                DataTable dataPercent = GetDataPercent();
                //chart = Worksheet.Drawings.AddChart("ColumnStackedChart", eChartType.Line) as ExcelLineChart;  
                chart = worksheet.Drawings.AddChart("ColumnStackedChart", eChartType.ColumnStacked) as ExcelBarChart;//设置图表样式  
                //chart.Legend.Position = eLegendPosition.Top;
                //chart.Legend.Add();
                chart.Legend.Remove();
                chart.Title.Text = reportTitle;//设置图表的名称  
                //chart.SetPosition(200, 50);//设置图表位置  
                chart.SetSize(800, 400);//设置图表大小  
                chart.ShowHiddenData = true;
               // chart.PlotArea.Fill.Color = Color.Red;
              //  chart.Fill.Color = Color.DarkBlue;
                //chart.YAxis.MinorUnit = 1;  
                chart.XAxis.MinorUnit = 1;//设置X轴的最小刻度  
                chart.DataLabel.ShowValue = true;
             
                //chart.YAxis.LabelPosition = eTickLabelPosition.High;
               // chart.YAxis.TickLabelPosition = eTickLabelPosition.NextTo;
                //chart.DataLabel.ShowCategory = true;  
                //chart.DataLabel.ShowPercent = true;//显示百分比  
               
                //设置月份  
                for (int col = 1; col <= dataPercent.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dataPercent.Columns[col - 1].ColumnName;
                }
                //设置数据  
                for (int row = 1; row <= dataPercent.Rows.Count; row++)
                {
                    for (int col = 1; col <= dataPercent.Columns.Count; col++)
                    {
                        string strValue = dataPercent.Rows[row - 1][col - 1].ToString();
                        if (col == 1)
                        {
                            worksheet.Cells[row + 1, col].Value = strValue;
                        }
                        else
                        {
                            double realValue = double.Parse(strValue);
                            worksheet.Cells[row + 1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                           // worksheet.Cells[row + 1, col].Style.Numberformat.Format = "#0\\.00%";//设置数据的格式为百分比  
                            worksheet.Cells[row + 1, col].Value = realValue;
                            worksheet.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            //if (realValue < 0.90d)//如果小于90%则该单元格底色显示为红色  
                            //{

                            //    worksheet.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.Red);
                            //}
                            //else if (realValue >= 0.90d && realValue <= 0.95d)//如果在90%与95%之间则该单元格底色显示为黄色  
                            //{
                            //    worksheet.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                            //}
                            //else
                            //{
                            //    worksheet.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.Green);//如果大于95%则该单元格底色显示为绿色  
                            //}
                        }
                    }
                    //chartSerie = chart.Series.Add(worksheet.Cells["A2:M2"], worksheet.Cells["B1:M1"]);  
                    //chartSerie.HeaderAddress = worksheet.Cells["A2"];  
                    //chart.Series.Add()方法所需参数为：chart.Series.Add(X轴数据区,Y轴数据区)  
                    chartSerie = chart.Series.Add(worksheet.Cells[row + 1, 2, row + 1, 2 + dataPercent.Columns.Count - 2], worksheet.Cells[row , 2, row , 2 + dataPercent.Columns.Count - 2]);

                    chartSerie.HeaderAddress = worksheet.Cells[row + 1, 1];//设置每条线的名称
                    
                 }
                //因为假定每家公司至少完成了80%以上，所以这里设置Y轴的最小刻度为80%，这样使图表上的折线更清晰  
                 //chart.YAxis.MinValue = 0.8d;
                
                //chart.SetPosition(200, 50);//可以通过制定左上角坐标来设置图表位置  
                //通过指定图表左上角所在的行和列及对应偏移来指定图表位置  
                //这里CommpanyNames.Length + 1及3分别表示行和列  
                  chart.SetPosition(CommpanyNames.Length + 1, 0, 3, 0);
                  //chart.Border.Fill.Color = Color.Yellow;
                

                #endregion research
                package.Save();//保存文件  
            }  
        
        }
        private static  DataTable GetDataPercent()
        {
           DataTable data = new  DataTable();
            DataRow row = null;
            Random random = new Random();

            data.Columns.Add(new DataColumn("公司名", typeof(string)));
            foreach (string monthName in MonthNames)
            {
                data.Columns.Add(new DataColumn(monthName, typeof(double)));
            }
            //每个公司每月的百分比表示完成的业绩与计划的百分比  
            for (int i = 0; i < CommpanyNames.Length; i++)
            {
                row = data.NewRow();
                row[0] = CommpanyNames[i];
                for (int j = 1; j <= MonthNames.Length; j++)
                {
                    //这里采用了随机生成数据，但假定每家公司至少完成了计划的85%以上  
                    row[j] =  random.Next(0, 1005);
                }
                data.Rows.Add(row);
            }


            return data;
        }

   }
}
