using System; 
using System.Linq;
using System.Text; 
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO; 
using Novacode;
using Svg;
using System.Drawing.Imaging;
using System.Drawing;
using System.Collections.Generic;


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

        public void CreateExcelXStream(List<SvgDocument> svgDocs, Stream stream)
        {
            using (stream)
            {
                IWorkbook workbook = new XSSFWorkbook();
         

                //for (int i = 0; i < svgDocs.Count; i++)
                //{
                //    using (MemoryStream ms = new MemoryStream())
                //    {
                //        System.Drawing.Bitmap image = svgDocs[i].Draw();
                //        image.Save(ms, ImageFormat.Bmp);
                //        ms.Seek(0, SeekOrigin.Begin);
                XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3); 
                       // workbook.AddPicture(ms.ToArray(), PictureType.DIB);
                int index1= LoadImage("c:\\1.png", workbook);
                int index2=LoadImage("c:\\2.png", workbook);
                ISheet sheet = workbook.CreateSheet();
                IDrawing patriarch = sheet.CreateDrawingPatriarch();
                        XSSFPicture pic1 = (XSSFPicture)(XSSFPicture)patriarch.CreatePicture(anchor,index1);
                        anchor = new XSSFClientAnchor(0, 0, 1023, 0, 10, 10, 11, 13);  
                        XSSFPicture pic2 = (XSSFPicture)(XSSFPicture)patriarch.CreatePicture(anchor, index2);
                        pic1.Resize();
                        pic2.Resize();
                //    }
                //}
                //workbook.Write(stream);
                        FileStream sw = File.Create("test.xlsx");
                        workbook.Write(sw);
                        sw.Close();
            }

        }

        public static int LoadImage(string path, IWorkbook wb)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            return wb.AddPicture(buffer, PictureType.JPEG);

        }
    }
}