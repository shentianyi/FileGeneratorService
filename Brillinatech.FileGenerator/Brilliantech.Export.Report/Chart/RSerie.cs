using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using System.Xml;
using Brilliantech.Export.Report.Enum;

namespace Brilliantech.Export.Report.Chart
{
    public class RSerie
    {
      //  private eChartType type=eChartType.ColumnStacked;
        private Color color=Color.LemonChiffon;
        private string colorString=null;

        //private int xStartRow;
        //private int xStartCol;
        //private int xEndRow;
        //private int xEndCol;

        //private int yStartRow;
        //private int yStartCol;
        //private int yEndRow;
        //private int yEndCol;
        private string xAixs;
        private string yAixs;
        private bool showDataLabel = true;

        private bool useSecondaryAxis = false;
        private ChartAxisFormatType yAxisFormatType = ChartAxisFormatType.None;
        private string headerAddress = null;
        public RSerie() { 
         
        }

        public RSerie(XmlNode parent) {            
             //string _type = parent.Attributes["type"].Value;
            //if (string.Equals(_type, "column"))
            //{
            //    this.type = eChartType.ColumnStacked;
            //}
            //else if(string.Equals(_type, "line")) {
            //    this.type = eChartType.Line;
            //}
            XmlElement ele=(XmlElement)parent;
            if (ele.HasAttribute("show_data_label")) {
                this.showDataLabel = bool.Parse(parent.Attributes["show_data_label"].Value);
            }

            if (ele.HasAttribute("color"))
            {
                this.colorString = parent.Attributes["color"].Value;
                if (this.colorString.StartsWith("#"))
                {
                    this.color = ColorTranslator.FromHtml(parent.Attributes["color"].Value);
                }
                else
                {
                    this.color = ColorTranslator.FromHtml("#" + parent.Attributes["color"].Value);
                }
            }

           // var d = parent.SelectSingleNode("xstart_row");
            this.xAixs = parent.SelectSingleNode("xaixs").FirstChild.Value;

            this.yAixs = parent.SelectSingleNode("yaixs").FirstChild.Value; 
            if (ele.HasAttribute("use_secondary_axis"))
            {
                this.useSecondaryAxis = bool.Parse(parent.Attributes["use_secondary_axis"].Value);
            }
            if (ele.HasAttribute("yaxis_format_type"))
            {
                this.yAxisFormatType = (ChartAxisFormatType)int.Parse(parent.Attributes["yaxis_format_type"].Value);
            }
            if (ele.HasAttribute("header_address")) {
                this.headerAddress = parent.Attributes["header_address"].Value;
            }
        }
        //public eChartType Type
        //{
        //    get { return type; }
        //}
        public Color Color
        {
            get { return color; }
            set { color = value; }
        }

        public string ColorString
        {
            get { return colorString; }
        }
        public string XAixs
        {
            get { return xAixs; }
            set { xAixs = value; }
        }
        public string YAixs
        {
            get { return yAixs; }
            set { yAixs = value; }
        }
        public bool ShowDataLabel
        {
            get { return showDataLabel; }
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

        public string YAxisFormatString
        {
            get
            {
                return ChartAxisFormat.GetFormatString(this.yAxisFormatType);
            }
        }
        public string HeaderAddress
        {
            get { return headerAddress; }
            set { headerAddress = value; }
        }

    }
}
