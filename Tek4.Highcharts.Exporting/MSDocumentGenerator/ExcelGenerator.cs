using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Svg;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Drawing.Imaging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Tek4.Highcharts.Exporting.Model;
using NPOI.HSSF.Util;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Tek4.Highcharts.Exporting.MSDocumentGenerator
{
    public class ExcelGenerator
    { 
        /// <summary>
        /// create xls
        /// using NPOI
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public static void CreateExcelStreamBySvgs(List<SvgDocument> svgDocs, Stream stream)
        {
            using (stream)
            {
                IWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");
                IDrawing patriarch = sheet.CreateDrawingPatriarch();
                HSSFClientAnchor anchor;
                IPicture pic;
                IRow row = null;
                for (int i = 0; i < svgDocs.Count; i++)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (System.Drawing.Bitmap image = svgDocs[i].Draw())
                        {
                            image.Save(ms, ImageFormat.Bmp);
                            ms.Seek(0, SeekOrigin.Begin);
                            int index = workbook.AddPicture(ms.ToArray(), PictureType.JPEG);
                            row = sheet.CreateRow(i);
                            row.HeightInPoints = convertPixelToPoints(image.Height) + 10;
                            anchor = new HSSFClientAnchor(0, 0, 0, 0, 0, i, 0, i);
                            pic = patriarch.CreatePicture(anchor, index);
                            pic.Resize();
                        }
                    }
                }
                workbook.Write(stream);
            }
        }
        /// <summary>
        /// create xls with table
        /// using NPOI
        /// </summary>
        /// <param name="svgDoc"></param>
        /// <param name="stream"></param>
        /// <param name="table"></param>
        public static void CreateExcelWithTableStreamBySvg(SvgDocument svgDoc, Stream stream, FileTable table) 
        {
            using (stream)
            {
                IWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Sheet1");
                IDrawing patriarch = sheet.CreateDrawingPatriarch();
                HSSFClientAnchor anchor;
                IPicture pic;

                ICellStyle head_style = createHeadCellStyle(workbook);
                ICellStyle cell_style = createCellBorderStyle(workbook);
                IRow head_row = null;
                head_row=sheet.CreateRow(0);
                ICell blank = head_row.CreateCell(0); 
                blank.CellStyle = head_style;

                for (int i = 0; i < table.series.Count; i++)
                {
                    ICell head_cell = head_row.CreateCell(i + 1);
                    head_cell.SetCellValue(table.series[i]);
                    head_cell.CellStyle = head_style;
                }
                // create vertical yAxis 
                int _ycount = 1;
                foreach (string key in table.xAxis.Keys)
                {
                    IRow y_row = sheet.CreateRow(_ycount);
                    ICell cell =y_row.CreateCell(0);
                    cell.SetCellValue(table.xAxis[key]);
                    cell.CellStyle = head_style;
                    for (int i = 0; i < table.series.Count; i++)
                    {
                        ICell value_cell = y_row.CreateCell(i + 1); 
                        value_cell.SetCellValue(table.rows[i][key]);
                        value_cell.CellStyle = cell_style;
                    }
                    _ycount++;
                } 
                using (MemoryStream ms = new MemoryStream())
                { 
                    using (System.Drawing.Bitmap image = svgDoc.Draw())
                    {
                        image.Save(ms, ImageFormat.Bmp);
                        ms.Seek(0, SeekOrigin.Begin);
                        int index = workbook.AddPicture(ms.ToArray(), PictureType.JPEG); 
                        anchor = new HSSFClientAnchor(0, 0, 0, 0, table.series.Count + 2, 0, 100, 100);
                        pic = patriarch.CreatePicture(anchor, index);
                        pic.Resize();
                    }
                }
                workbook.Write(stream);
            }
        }

        /// <summary>
        /// create xlsx
        /// using EPPlus
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public static void CreateExcelXStreamBySvgs(List<SvgDocument> svgDocs, Stream stream)
        {
            using (stream)
            {
                ExcelPackage package = new ExcelPackage();
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
                ExcelPicture picture;

                package.DoAdjustDrawings = false;
                for (int i = 0; i < svgDocs.Count; i++)
                {
                    using (System.Drawing.Bitmap image = svgDocs[i].Draw())
                    {
                        picture = sheet.Drawings.AddPicture(i.ToString(), image);
                        picture.SetPosition(i, 0, 0, 0);
                        sheet.Row(i + 1).Height = convertPixelToPoints(image.Height) + 10;
                        picture.SetSize(100);
                    }
                }
                package.SaveAs(stream);
            }
        }

        /// <summary>
        /// create xlsx with table
        /// using EPPlus
        /// </summary>
        /// <param name="svgDoc"></param>
        /// <param name="stream"></param>
        /// <param name="table"></param>
        public static void CreateExcelXWithTableStreamBySvg(SvgDocument svgDoc, Stream stream, FileTable table)
        {
            using (stream)
            {
                ExcelPackage package = new ExcelPackage();
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
                ExcelPicture picture;

                setHeadCellStyle(sheet.Cells[1, 1]);
                for (int i = 0; i < table.series.Count; i++)
                {
                    sheet.Cells[1, i + 2].Value = table.series[i];
                    setHeadCellStyle(sheet.Cells[1, i + 2]);
                }

                int _ycount = 2;
                foreach (string key in table.xAxis.Keys)
                {
                    sheet.Cells[_ycount, 1].Value = table.xAxis[key];
                    setHeadCellStyle(sheet.Cells[_ycount, 1]);
                    for (int i = 0; i < table.series.Count; i++)
                    {
                        sheet.Cells[_ycount, i + 2].Value = table.rows[i][key];
                        setCellBorderStyle(sheet.Cells[_ycount, i + 2]);
                    }
                    _ycount++;
                }

                using (System.Drawing.Bitmap image = svgDoc.Draw())
                {
                    picture = sheet.Drawings.AddPicture("0", image);
                    picture.SetPosition(0, 0, table.series.Count+2, 0);
                    picture.SetSize(100);
                }
                package.SaveAs(stream);
            }
        }

        /// <summary>
        /// convert pixel to point
        /// </summary>
        /// <param name="pixel"></param>
        /// <returns></returns>
        private static float convertPixelToPoints(int pixel)
        {
            return pixel * 72 / 96;
        }

            /// <summary>
        /// create head cell style
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private static ICellStyle createHeadCellStyle(IWorkbook workbook)
        {
            // insert table
            ICellStyle style = createCellBorderStyle(workbook); 
            style.FillForegroundColor = HSSFColor.LIGHT_GREEN.index;
            style.FillPattern = FillPatternType.SOLID_FOREGROUND;
            IFont font = workbook.CreateFont();
            font.Boldweight =(short) FontBoldWeight.BOLD;
            font.FontName = "Calibri";
            font.FontHeightInPoints = 11;
            style.SetFont(font);
            return style;
        }

        /// <summary>
        /// create value cell style
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private static ICellStyle createCellBorderStyle(IWorkbook workbook)
        {
            // insert table
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.THIN;
            style.BorderLeft = BorderStyle.THIN;
            style.BorderRight = BorderStyle.THIN;
            style.BorderTop = BorderStyle.THIN; 
            IFont font = workbook.CreateFont(); 
            font.FontName = "Calibri";
            font.FontHeightInPoints = 11;
            style.SetFont(font);
            return style;
        }

        /// <summary>
        /// set head cell style
        /// </summary>
        /// <param name="cellRange"></param>
        private static void setHeadCellStyle(ExcelRange cellRange) {
            cellRange.Style.Font.Bold = true;
            cellRange.Style.Border.Top.Style=cellRange.Style.Border.Right.Style=cellRange.Style.Border.Bottom.Style=cellRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cellRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellRange.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
        }

        /// <summary>
        /// set value cell style
        /// </summary>
        /// <param name="cellRange"></param>
        private static void setCellBorderStyle(ExcelRange cellRange)
        {
            cellRange.Style.Border.Top.Style = cellRange.Style.Border.Right.Style = cellRange.Style.Border.Bottom.Style = cellRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }
    }
}
