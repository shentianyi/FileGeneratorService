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
               // HSSFWorkbook hssfworkbook = new HSSFWorkbook();
               // //将图片加入Workbook
               // byte[] bytes = System.IO.File.ReadAllBytes(@"c:\\1.png");
               // int pictureIdx1 = hssfworkbook.AddPicture(bytes, PictureType.JPEG);

               // bytes = System.IO.File.ReadAllBytes(@"c:\\2.png");
               // int pictureIdx2 = hssfworkbook.AddPicture(bytes, PictureType.JPEG);

               // //获取存在的Sheet，必须在AddPicture之后
               //ISheet sheet = hssfworkbook.CreateSheet("Sheet1");
               // IDrawing patriarch = sheet.CreateDrawingPatriarch();

               // //插入图片
               // HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
               // IPicture pict1 = patriarch.CreatePicture(anchor, pictureIdx1);
               // pict1.Resize();
               // anchor = new HSSFClientAnchor(0, 0, 1023, 0, 10, 10, 11, 13);
               // IPicture pict2 = patriarch.CreatePicture(anchor, pictureIdx2);
               // pict2.Resize();
               // FileStream fs = File.Create("test.xls");
               // hssfworkbook.Write(fs);
               // fs.Close();
                IWorkbook workbook = new XSSFWorkbook();
                 
                byte[] bytes = File.ReadAllBytes(@"c:\\1.png");
                int pic1 = workbook.AddPicture(bytes, PictureType.JPEG);
                bytes = File.ReadAllBytes(@"c:\\2.png");
                int pic2 = workbook.AddPicture(bytes, PictureType.JPEG);
                ISheet sheet = workbook.CreateSheet("Sheet1");
                IDrawing pariarch = sheet.CreateDrawingPatriarch();
                XSSFClientAnchor anchor = new XSSFClientAnchor();
                IPicture pict1 = pariarch.CreatePicture(anchor, pic1);
                pict1.Resize();

                IPicture pict2 = pariarch.CreatePicture(anchor, pic2);
                pict2.Resize();
               // workbook.Write(stream); 
                ////for (int i = 0; i < svgDocs.Count; i++)
                ////{
                ////    using (MemoryStream ms = new MemoryStream())
                ////    {
                ////        System.Drawing.Bitmap image = svgDocs[i].Draw();
                ////        image.Save(ms, ImageFormat.Bmp);
                ////        ms.Seek(0, SeekOrigin.Begin);
                //XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
                //// workbook.AddPicture(ms.ToArray(), PictureType.DIB);
                //int index1 = LoadImage("c:\\1.png", workbook);

                //int index2 = LoadImage("c:\\2.png", workbook);

                //ISheet sheet = workbook.CreateSheet();
                //IDrawing patriarch = sheet.CreateDrawingPatriarch();
                //XSSFPicture pic1 =(XSSFPicture)patriarch.CreatePicture(anchor, index1);
                //XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, 1023, 0, 10, 10, 11, 13);
                //XSSFPicture pic2 =(XSSFPicture)patriarch.CreatePicture(anchor2, index2);
                //pic1.Resize();
                //pic2.Resize();
                ////    }
                ////}
                ////workbook.Write(stream);
                FileStream sw = File.Create("test.xlsx");
                workbook.Write(sw);
                sw.Close();

            }
        }
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
                        System.Drawing.Bitmap image = svgDocs[i].Draw();
                        image.Save(ms, ImageFormat.Bmp);
                        ms.Seek(0, SeekOrigin.Begin);
                        int index = workbook.AddPicture(ms.ToArray(), PictureType.JPEG);
                        row = sheet.CreateRow(i);
                        row.HeightInPoints =  image.Height;
                        anchor = new HSSFClientAnchor(0, 0, 0, 0, 0, i, 0, i);
                        pic = patriarch.CreatePicture(anchor, index);
                        pic.Resize();
                    }
                }
                workbook.Write(stream);
            }
        }

    }
}