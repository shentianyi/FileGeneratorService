using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace Brilliantech.Export.Report.Chart
{
    public class RChart
    {
        private string title;
        private int? height;
        private int? width;
        private bool showLegend = false;
        private RSerie[] series;

        public RChart() { }

        public RChart(XmlNode parent)
        {
            title = parent.Attributes["title"].Value;
            XmlElement ele=(XmlElement) parent;

            if (ele.HasAttribute("height"))
            {
                height = int.Parse(parent.Attributes["height"].Value);
            } 
            if (ele.HasAttribute("width") )
            {
                width = int.Parse(parent.Attributes["width"].Value);
            }
            if (ele.HasAttribute("show_legend"))
            {
                showLegend = bool.Parse(parent.Attributes["show_legend"].Value);
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

        public string Title
        {
            get { return title; } 
        }
        public int? Height
        {
            get { return height; } 
        }
        public int? Width
        {
            get { return width; } 
        }
        public RSerie[] Series
        {
            get { return series; } 
        }
        public bool ShowLegend
        {
            get { return showLegend; }
        }
    }
}
