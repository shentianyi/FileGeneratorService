using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenExcel.OfficeOpenXml.Style;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TestConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
           // Console.WriteLine((BorderStyleValues)(int)ExcelBorderStyleValues.DashDot);
          //OpenXMLTest.Run();
          // OpenXMLTest.RunMerge();
           // OpenXMLChartTest.Run();
            EPPlusExcelChartTest.Run();
         //   BTReportChartTest.Run();
           // Console.Read();
        }
    }
}
