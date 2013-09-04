using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;


namespace Tek4.ConsoleTip
{
    class Program
    {
        static void Main(string[] args)
        {

            string filePath = "d:\\a.pptx";
          // new PowerPointCreator().CreatePresentation(filePath);
           // Console.Read();
           new PowerPointPicInstor().Insert(filePath);
        }

       
    }
}
