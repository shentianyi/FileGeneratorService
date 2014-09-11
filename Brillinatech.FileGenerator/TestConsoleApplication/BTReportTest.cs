using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Brilliantech.Export.Report;

namespace TestConsoleApplication
{
    public class BTReportTest
    {
        public static void Run() {
            string xml;
            
            using (FileStream file = new FileStream("data\\report_meta.xml", FileMode.Open, FileAccess.Read))
            {
                using (MemoryStream ms = new MemoryStream()) {
                    file.CopyTo(ms); 
                    StreamReader sr = new StreamReader(ms);
                    ms.Position = 0;
                    xml = sr.ReadToEnd();
                }    
            }

            IReport r = new ReportTable("C:\\Excel\\BT"+DateTime.Now.ToString("yyyy-MM-ddHHmmsss")+".xlsx",xml);
            r.Generate();
        }
    }
}