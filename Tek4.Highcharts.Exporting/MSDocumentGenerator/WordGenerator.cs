using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Svg;
using System.IO;
using System.Drawing.Imaging;
using Novacode;

namespace Tek4.Highcharts.Exporting.MSDocumentGenerator
{
   public class WordGenerator
   { /// <summary>
       /// create doc or docx 
       /// using DOCX: Novacode
       /// </summary>
       public static void CreateDocStreamBySvgs(List<SvgDocument> svgDocs, Stream stream)
       {
           using (stream)
           {
               using (DocX doc = DocX.Create(stream))
               {
                   Novacode.Paragraph p = doc.InsertParagraph("", false);
                   for (int i = 0; i < svgDocs.Count; i++)
                   {
                       using (MemoryStream ms = new MemoryStream())
                       {
                           System.Drawing.Bitmap image = svgDocs[i].Draw();
                           image.Save(ms, ImageFormat.Bmp);
                           ms.Seek(0, SeekOrigin.Begin);
                           Novacode.Image img = doc.AddImage(ms);
                           Novacode.Picture pic = img.CreatePicture();
                           p.AppendPicture(pic);
                       }
                   }
                   doc.Save();
               }
           }
       }
       
    }
}
