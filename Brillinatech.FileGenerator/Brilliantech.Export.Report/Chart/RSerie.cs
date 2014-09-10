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
        private int startRow;
        private int startCol;
        private int endRow;
        private int endCol;

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

        public int StartRow
        {
            get { return startRow; }
            set { startRow = value; }
        }
        public int StartCol
        {
            get { return startCol; }
            set { startCol = value; }
        }
        public int EndRow
        {
            get { return endRow; }
            set { endRow = value; }
        }
        public int EndCol
        {
            get { return endCol; }
            set { endCol = value; }
        }
    }
}
