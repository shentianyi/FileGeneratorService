using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;
using Brilliantech.Export.Report.Enum;

namespace Brilliantech.Export.Report.Chart
{
    public class RChart
    {
        private string title;
        private int? height;
        private int? width;
        private bool showLegend = false;
        private bool useSecondaryAxis = false;
        private ChartAxisFormatType yAxisFormatType = ChartAxisFormatType.None;

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

            if (ele.HasAttribute("use_secondary_axis"))
            {
                this.useSecondaryAxis = bool.Parse(parent.Attributes["use_secondary_axis"].Value);
            }
            if (ele.HasAttribute("yaxis_format_type"))
            {
                this.yAxisFormatType = (ChartAxisFormatType)int.Parse(parent.Attributes["yaxis_format_type"].Value);
            }
        }

        public string YAxisFormatString
        {
            get
            {
                return ChartAxisFormat.GetFormatString(this.yAxisFormatType);
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
        public bool UseSecondaryAxis
        {
            get { return useSecondaryAxis; }
        }
         
        public ChartAxisFormatType AxisFormatType
        {
            get { return yAxisFormatType; }
            set { yAxisFormatType = value; }
        }
    }
}
