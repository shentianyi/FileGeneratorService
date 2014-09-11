using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.IO;
using Brilliantech.Export.Report.XmlParser;
using Brilliantech.Export.Report.Table;  

namespace Brilliantech.Export.Report
{
    public class Report
    {
        protected string filePath;
        protected FileInfo fileInfo;
        protected string xml;
        protected TableXmlParser tableXmlParser;
        protected ExcelPackage package;
        protected ExcelWorksheet worksheet;
        protected RTColumn[][] columns;


        protected int cols_stat;
        protected int rows_stat;
        protected int tableWidth = 0;
      //  private double rowHeight = 22.5;
        protected double defaultRowHeight = 15.0;

        public int rowOffsetCount = 0;
        public int colCount = 0;

        public Report() { }

        //public Report(string filePath, string xml)
        //{
        //    this.filePath = filePath;
        //    this.fileInfo = new FileInfo(filePath);
        //    this.xml = xml;
        //    tableXmlParser = new TableXmlParser(xml);
        //}

        //public void Generate()
        //{
        //    using (package = new ExcelPackage(this.FileInfo))
        //    {
        //        worksheet = package.Workbook.Worksheets.Add("Sheet1");
        //        GenerateTableHead();
        //        MergeTableHead();
        //        GenerateTableBody();
        //        package.Save();
        //    }
        //}

        /// <summary>
        /// generate head
        /// </summary>
        protected void GenerateTableHead()
        {
            columns = tableXmlParser.GetColumnsInfo();
            int[] widths = tableXmlParser.Widths;
            this.cols_stat = widths.Length;

            for (int i = 0; i < widths.Length; i++)
            {
                tableWidth += widths[i];
            }
            for (int row = 1; row <= columns.Length; row++)
            {
              //  worksheet.Row(row).Height = rowHeight;
                for (int col = 1; col <= columns[row - 1].Length; col++)
                {
                    worksheet.Cells[row, col].Value = columns[row - 1][col - 1].ColName;                
                    worksheet.Column(col).Width = widths[col - 1]/6;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[row, col].Style.Border.BorderAround(columns[row - 1][col - 1].BorderStyle, columns[row - 1][col - 1].BorderColor);
                    worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(columns[row - 1][col - 1].BackgroundColor);
                    
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                    worksheet.Cells[row, col].Style.Font.Name = "Arial";
                    worksheet.Cells[row, col].Style.Font.Size = 10;
                }
            }
            rowOffsetCount = columns.Length;
        }

        /// <summary>
        /// merge table head
        /// </summary>
        protected void MergeTableHead()
        {
            for (int row = 1; row <= columns.Length; row++)
            {
                int row_index = row - 1;
                for (int col = 1; col <= columns[row_index].Length; col++)
                {
                    int colspan = columns[row_index][col - 1].Colspan;
                    if (colspan > 0)
                    {
                        worksheet.Cells[row, col, row, col + colspan - 1].Merge = true;
                    }

                    int rowspan = columns[row_index][col - 1].Rowspan;
                    if (rowspan > 0)
                    {
                        worksheet.Cells[row, col, row + rowspan - 1, col].Merge = true;
                    }
                }
            }
        }

        /// <summary>
        /// generate body
        /// </summary>
        protected void GenerateTableBody()
        {
            RTRow[] rows = tableXmlParser.GetRows();
            this.rows_stat = rows.Length;
            for (int row = 1; row <= rows.Length; row++) {
                RTCell[] cells = rows[row - 1].Cells;
                int rowNum = row + rowOffsetCount;
            //    worksheet.Row(rowNum).Height = rowHeight;

                for (int col = 1; col <= cells.Length; col++) {
                    worksheet.Cells[rowNum, col].Value = cells[col - 1].GetValue();

                    worksheet.Cells[rowNum, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[rowNum, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[rowNum, col].Style.Border.BorderAround(cells[col-1].BorderStyle,cells[col-1].BorderColor);
                    worksheet.Cells[rowNum, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[rowNum, col].Style.Fill.BackgroundColor.SetColor(RTCell.GetBackgroundColor(row - 1));

                    worksheet.Cells[rowNum, col].Style.Font.Name = "Arial";
                    worksheet.Cells[rowNum, col].Style.Font.Size = 10;
                }
            }
            rowOffsetCount += rows.Length;
        }
        public FileInfo FileInfo
        {
            get { return fileInfo; }
            set { fileInfo = value; }
        }

        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }
    }
}
