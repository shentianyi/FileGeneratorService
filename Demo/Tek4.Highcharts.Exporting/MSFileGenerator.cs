using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using NPOI.XWPF.UserModel;
using System.IO;
using System.Drawing;
using Novacode;


namespace Tek4.Highcharts.Exporting
{
    public class MSFileGenerator
    {
        /// <summary>
        /// create doc or docx 
        /// using DOCX: Novacode
        /// </summary>
        public void CreateDoc()
        {
            //    DocX document = DocX.Create("doc.doc");
            //    document.AddImage(@"tmp\identicon.png");
            //    document.Save();
            using (DocX doc = DocX.Create("a.doc"))
            {
                using (MemoryStream ms = new MemoryStream()) {
                    System.Drawing.Image image = System.Drawing.Image.FromFile("c:\\identicon.png");
                    image.Save(ms, image.RawFormat);
                    ms.Seek(0, SeekOrigin.Begin);
                    Novacode.Image img = doc.AddImage(ms);
                    Paragraph p = doc.InsertParagraph("",false);
                    Picture pic = img.CreatePicture();
                    p.InsertPicture(pic);
                    doc.Save();
                }
            }
        }
    }
}