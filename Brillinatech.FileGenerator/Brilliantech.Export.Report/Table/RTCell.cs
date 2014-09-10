using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace Brilliantech.Export.Report.Table
{
    public class RTCell
    {
        private string value = "";
        private Color backgroundColor = ColorTranslator.FromHtml("#E3EFFF");
        private ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin;
        private Color borderColor = ColorTranslator.FromHtml("#A4BED4");

        public RTCell() { }

        public RTCell(XmlNode parent)
        {
            if (parent.HasChildNodes)
                value = parent.FirstChild.Value;
            XmlElement el = (XmlElement)parent; 
        }


        public string Value
        {
            get { return this.value; }
        }

        public Color BackgroundColor
        {
            get { return backgroundColor; }
            set { backgroundColor = value; }
        }

        public ExcelBorderStyle BorderStyle
        {
            get { return borderStyle; }
            set { borderStyle = value; }
        }
        public Color BorderColor
        {
            get { return borderColor; }
            set { borderColor = value; }
        }

        public static Color GetBackgroundColor(int row)
        {
            return row % 2 == 0 ? ColorTranslator.FromHtml("#FFFFFF") : ColorTranslator.FromHtml("#E3EFFF");
        }

        public object GetValue()
        {
            int iVal;
            double dVal;
            if (int.TryParse(this.value, out iVal))
            {
                return iVal;
            }
            else if (double.TryParse(this.value, out dVal))
            {
                return dVal;
            }
            return this.value;
        }
    }
}
