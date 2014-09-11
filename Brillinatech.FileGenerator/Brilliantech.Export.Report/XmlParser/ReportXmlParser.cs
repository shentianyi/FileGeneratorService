using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Brilliantech.Export.Report.XmlParser
{
    public class ReportXmlParser
    {
        protected XmlElement root;
       
        public ReportXmlParser() { }


        public ReportXmlParser(string xml)
        {
            XmlDocument dom = new XmlDocument();
            try
            {
                dom.LoadXml(xml);
            }
            catch (Exception e)
            {

            }
            root = dom.DocumentElement;
        } 
    }
}
