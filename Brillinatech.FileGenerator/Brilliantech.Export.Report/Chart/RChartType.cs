using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace Brilliantech.Export.Report.Chart
{
    public class RChartType
    {
        private eChartType type = eChartType.ColumnStacked;
        private RSerie[] series;

        public RChartType(XmlNode parent)
        {

            string _type = parent.Attributes["type"].Value;
            if (string.Equals(_type, "column"))
            {
                this.type = eChartType.ColumnStacked;
            }
            else if (string.Equals(_type, "line"))
            {
                this.type = eChartType.Line;
            } 
            XmlNodeList nodes = ((XmlElement)parent).GetElementsByTagName("serie");
            if (nodes != null && nodes.Count > 0)
            {
                this.series = new RSerie[nodes.Count];
                for (int i = 0; i < nodes.Count; i++)
                {
                    series[i] = new RSerie(nodes[i]);
                }
            }
        }
        public eChartType Type
        {
            get { return type; }
        }
        public RSerie[] Series
        {
            get { return series; }
        }
    }
}
