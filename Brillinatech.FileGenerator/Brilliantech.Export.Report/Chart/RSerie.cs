using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using System.Xml;

namespace Brilliantech.Export.Report.Chart
{
    public class RSerie
    {
        private eChartType type=eChartType.ColumnStacked;
        private Color color=Color.LemonChiffon;
        private int xStartRow;
        private int xStartCol;
        private int xEndRow;
        private int xEndCol;

        private int yStartRow;
        private int yStartCol;
        private int yEndRow;
        private int yEndCol;

        public RSerie() { 
         
        }

        public RSerie(XmlNode parent) { 
            
            string _type = parent.Attributes["type"].Value;
            if (string.Equals(_type, "column"))
            {
                this.type = eChartType.ColumnStacked;
            }
            else if(string.Equals(_type, "line")) {
                this.type = eChartType.Line;
            }
            this.color = ColorTranslator.FromHtml(parent.Attributes["color"].Value);
           // var d = parent.SelectSingleNode("xstart_row");
         
            this.xStartRow = int.Parse(parent.SelectSingleNode("xstart_row").FirstChild.Value);
            this.xStartCol = int.Parse(parent.SelectSingleNode("xstart_col").FirstChild.Value);
            this.xEndRow = int.Parse(parent.SelectSingleNode("xend_row").FirstChild.Value);
            this.xEndCol = int.Parse(parent.SelectSingleNode("xend_col").FirstChild.Value);

            this.yStartRow = int.Parse(parent.SelectSingleNode("ystart_row").FirstChild.Value);
            this.yStartCol = int.Parse(parent.SelectSingleNode("ystart_col").FirstChild.Value);
            this.yEndRow = int.Parse(parent.SelectSingleNode("yend_row").FirstChild.Value);
            this.yEndCol = int.Parse(parent.SelectSingleNode("yend_col").FirstChild.Value);
        }
        public eChartType Type
        {
            get { return type; }
        }
        public Color Color
        {
            get { return color; }
            set { color = value; }
        }

        public int XStartRow
        {
            get { return xStartRow; }
            set { xStartRow = value; }
        }
        public int XStartCol
        {
            get { return xStartCol; }
            set { xStartCol = value; }
        }
        public int XEndRow
        {
            get { return xEndRow; }
            set { xEndRow = value; }
        }
        public int XEndCol
        {
            get { return xEndCol; }
            set { xEndCol = value; }
        }

        public int YStartRow
        {
            get { return yStartRow; }
            set { yStartRow = value; }
        }
        public int YStartCol
        {
            get { return yStartCol; }
            set { yStartCol = value; }
        }
        public int YEndRow
        {
            get { return yEndRow; }
            set { yEndRow = value; }
        }
        public int YEndCol
        {
            get { return yEndCol; }
            set { yEndCol = value; }
        }
    }
}
