using System; 
using System.Linq;
using System.Text;  
using System.IO; 
using Novacode;
using Svg;
using System.Drawing.Imaging;
using System.Drawing;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing; 


namespace Tek4.Highcharts.Exporting
{
    public class MSFileGenerator
    {
        /// <summary>
        /// create doc or docx 
        /// using DOCX: Novacode
        /// </summary>
        public void CreateDocStream(List<SvgDocument> svgDocs, Stream stream)
        {
            using (stream)
            {
                using (DocX doc = DocX.Create(stream))
                {
                    Paragraph p = doc.InsertParagraph("", false);
                    for (int i = 0; i < svgDocs.Count; i++)
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            System.Drawing.Bitmap image = svgDocs[i].Draw();
                            image.Save(ms, ImageFormat.Bmp);
                            ms.Seek(0, SeekOrigin.Begin);
                            Novacode.Image img = doc.AddImage(ms);
                            Picture pic = img.CreatePicture();
                            p.AppendPicture(pic);
                        }
                    }
                    doc.Save();
                }
            }
        }
        /// <summary>
        /// create xlsx
        /// using EPPlus
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public void CreateExcelXStream(List<SvgDocument> svgDocs, Stream stream)
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
        /// create xls
        /// using NPOI
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public void CreateExcelStream(List<SvgDocument> svgDocs, Stream stream)
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

        private float convertPixelToPoints(int pixel) {
            return pixel * 72 / 96;
        }

    }
}